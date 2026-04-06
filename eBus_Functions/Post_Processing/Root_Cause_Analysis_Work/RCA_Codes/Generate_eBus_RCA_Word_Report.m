function reportOutput = Generate_eBus_RCA_Word_Report(resultsInput, outputFolder, options)
% Generate_eBus_RCA_Word_Report  Create Word reports for eBus RCA results.
%
% README
% 1. Usage:
%    reportOutput = Generate_eBus_RCA_Word_Report;
%    reportOutput = Generate_eBus_RCA_Word_Report(results);
%    reportOutput = Generate_eBus_RCA_Word_Report('C:\path\to\RCA_Results.mat');
%    reportOutput = Generate_eBus_RCA_Word_Report(results, 'C:\path\to\output');
%
% 2. Inputs:
%    - resultsInput may be:
%      a) empty: placeholder-driven template/sample generation
%      b) a results struct returned by Vehicle_Detailed_Analysis
%      c) a MAT-file path containing "results" or "RCA_Results"
%
% 3. Outputs:
%    - eBus_Simulation_Root_Cause_Analysis_Report.docx
%    - eBus_Simulation_Root_Cause_Analysis_Report_Template.docx
%    - eBus_Simulation_Root_Cause_Analysis_Report_Sample.docx
%
% 4. Dependency note:
%    This function uses Microsoft Word COM automation. It requires Windows
%    and a local Microsoft Word installation, but no MATLAB add-on toolbox.

if nargin < 1
    resultsInput = [];
end
if nargin < 2
    outputFolder = [];
end
if nargin < 3 || isempty(options)
    options = struct();
end

options = localNormalizeOptions(options);
[results, sourceInfo] = localResolveResults(resultsInput);
outputFolder = localResolveOutputFolder(results, outputFolder);
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

if ~ispc
    error('Generate_eBus_RCA_Word_Report:Platform', ...
        'Word report generation requires Windows because it uses Microsoft Word COM automation.');
end

reportData = localBuildReportData(results, sourceInfo, options);

wordApp = [];
try
    wordApp = actxserver('Word.Application');
    wordApp.Visible = false;
    wordApp.DisplayAlerts = 0;
catch wordException
    error('Generate_eBus_RCA_Word_Report:WordUnavailable', ...
        'Microsoft Word automation could not be started: %s', wordException.message);
end

reportPath = fullfile(outputFolder, localReportFileName(reportData.Options.Language));
templatePath = "";
samplePath = "";

try
    localCreateReportDocument(wordApp, reportData, reportPath, 'report');
    if reportData.Options.CreateSupportingTemplateFiles
        supportFolder = fullfile(outputFolder, 'supporting_templates');
        if ~exist(supportFolder, 'dir')
            mkdir(supportFolder);
        end
        templatePath = fullfile(supportFolder, localTemplateFileName(reportData.Options.Language));
        samplePath = fullfile(supportFolder, localSampleFileName(reportData.Options.Language));
        localCreateReportDocument(wordApp, reportData, templatePath, 'template');
        localCreateReportDocument(wordApp, reportData, samplePath, 'sample');
    end
catch reportException
    localCleanupWord(wordApp);
    rethrow(reportException);
end

localCleanupWord(wordApp);

reportOutput = struct();
reportOutput.OutputFolder = string(outputFolder);
reportOutput.ReportFile = string(reportPath);
reportOutput.TemplateFile = string(templatePath);
reportOutput.SampleFile = string(samplePath);
reportOutput.Source = sourceInfo;

fprintf('\nWord report generation completed.\n');
fprintf('  Report   : %s\n', reportPath);
if strlength(templatePath) > 0
    fprintf('  Template : %s\n', templatePath);
end

function fileName = localReportFileName(language)
if string(language) == "DE"
    fileName = 'eBus_Simulation_Root_Cause_Analysis_Report_DE.docx';
else
    fileName = 'eBus_Simulation_Root_Cause_Analysis_Report.docx';
end
end

function fileName = localTemplateFileName(language)
if string(language) == "DE"
    fileName = 'eBus_Simulation_Root_Cause_Analysis_Report_Template_DE.docx';
else
    fileName = 'eBus_Simulation_Root_Cause_Analysis_Report_Template.docx';
end
end

function fileName = localSampleFileName(language)
if string(language) == "DE"
    fileName = 'eBus_Simulation_Root_Cause_Analysis_Report_Sample_DE.docx';
else
    fileName = 'eBus_Simulation_Root_Cause_Analysis_Report_Sample.docx';
end
end
if strlength(samplePath) > 0
    fprintf('  Sample   : %s\n', samplePath);
end
end

function options = localNormalizeOptions(options)
defaults = struct();
defaults.Project = "eBus Program [Insert Program Name]";
defaults.Author = string(getenv('USERNAME'));
if strlength(defaults.Author) == 0
    defaults.Author = "Simulation Engineer [Insert Name]";
end
defaults.Version = "1.0";
defaults.Confidentiality = "Internal / Confidential";
defaults.Company = "Company / Department [Insert]";
defaults.IncludeAppendixTables = true;
defaults.MaxAppendixRows = 120;
defaults.MaxSummaryRows = 12;
defaults.MaxSubsystemFigures = 10;
defaults.TemplateSubtitle = "Vehicle-Level and Subsystem-Level Technical Assessment";
defaults.PlaceholderTag = "[Insert]";
defaults.DateString = string(datetime('now', 'Format', 'dd-MMM-yyyy'));
defaults.CreateSupportingTemplateFiles = false;
defaults.ActiveSubsystemReportScope = "ALL";
defaults.Language = "EN";

fields = fieldnames(defaults);
for iField = 1:numel(fields)
    if ~isfield(options, fields{iField}) || isempty(options.(fields{iField}))
        options.(fields{iField}) = defaults.(fields{iField});
    end
end
options.Language = localNormalizeLanguage(options.Language);
end

function language = localNormalizeLanguage(languageValue)
language = upper(strtrim(char(string(languageValue))));
if any(strcmp(language, {'DE', 'DEU', 'GERMAN', 'DEUTSCH'}))
    language = "DE";
else
    language = "EN";
end
end

function [results, sourceInfo] = localResolveResults(resultsInput)
results = [];
sourceInfo = struct('HasResults', false, 'SourceType', "PlaceholderOnly", 'SourcePath', "", 'Description', "No RCA result struct supplied.");

if isempty(resultsInput)
    [results, sourceInfo] = localResolveImplicitResults();
    return;
end

if isstruct(resultsInput)
    results = resultsInput;
    sourceInfo.HasResults = true;
    sourceInfo.SourceType = "Struct";
    sourceInfo.Description = "RCA results supplied directly from MATLAB workspace.";
    if isfield(resultsInput, 'AnalysisData') && isfield(resultsInput.AnalysisData, 'RawData') && isfield(resultsInput.AnalysisData.RawData, 'MatFilePath')
        sourceInfo.SourcePath = string(resultsInput.AnalysisData.RawData.MatFilePath);
    end
    return;
end

if ischar(resultsInput) || (isstring(resultsInput) && isscalar(resultsInput))
    matPath = char(string(resultsInput));
    if ~isfile(matPath)
        error('Generate_eBus_RCA_Word_Report:MissingInputFile', 'Input MAT file not found: %s', matPath);
    end
    loaded = load(matPath);
    if isfield(loaded, 'results')
        results = loaded.results;
    elseif isfield(loaded, 'RCA_Results')
        results = loaded.RCA_Results;
    else
        candidateNames = fieldnames(loaded);
        found = false;
        for iName = 1:numel(candidateNames)
            candidate = loaded.(candidateNames{iName});
            if isstruct(candidate) && isfield(candidate, 'VehicleKPI') && isfield(candidate, 'SubsystemResults')
                results = candidate;
                found = true;
                break;
            end
        end
        if ~found
            error('Generate_eBus_RCA_Word_Report:InvalidInputFile', ...
                'MAT file does not contain a recognizable RCA results struct: %s', matPath);
        end
    end
    sourceInfo.HasResults = true;
    sourceInfo.SourceType = "MATFile";
    sourceInfo.SourcePath = string(matPath);
    sourceInfo.Description = "RCA results loaded from MAT file.";
    return;
end

error('Generate_eBus_RCA_Word_Report:UnsupportedInput', ...
    'resultsInput must be empty, a results struct, or a MAT file path.');
end

function [results, sourceInfo] = localResolveImplicitResults()
results = [];
sourceInfo = struct('HasResults', false, 'SourceType', "PlaceholderOnly", 'SourcePath', "", 'Description', "No RCA result struct supplied.");

try
    if evalin('base', 'exist(''RCA_Results'',''var'')')
        baseResults = evalin('base', 'RCA_Results');
        if isstruct(baseResults) && isfield(baseResults, 'VehicleKPI') && isfield(baseResults, 'SubsystemResults')
            results = baseResults;
            sourceInfo.HasResults = true;
            sourceInfo.SourceType = "BaseWorkspace";
            sourceInfo.Description = "RCA results recovered from MATLAB base workspace variable RCA_Results.";
            if isfield(baseResults, 'AnalysisData') && isfield(baseResults.AnalysisData, 'RawData') && isfield(baseResults.AnalysisData.RawData, 'MatFilePath')
                sourceInfo.SourcePath = string(baseResults.AnalysisData.RawData.MatFilePath);
            end
            return;
        end
    end
catch
end

try
    config = RCA_Config();
    defaultResultsRoot = fullfile(fileparts(mfilename('fullpath')), config.General.DefaultResultsFolder);
    resultsFiles = dir(fullfile(defaultResultsRoot, '**', 'RCA_Results.mat'));
    if ~isempty(resultsFiles)
        [~, order] = sort([resultsFiles.datenum], 'descend');
        latestPath = fullfile(resultsFiles(order(1)).folder, resultsFiles(order(1)).name);
        loaded = load(latestPath);
        if isfield(loaded, 'results')
            results = loaded.results;
        elseif isfield(loaded, 'RCA_Results')
            results = loaded.RCA_Results;
        end
        if ~isempty(results)
            sourceInfo.HasResults = true;
            sourceInfo.SourceType = "LatestSavedMAT";
            sourceInfo.SourcePath = string(latestPath);
            sourceInfo.Description = "RCA results loaded from the most recent RCA_Results.mat file.";
            return;
        end
    end
catch
end
end

function outputFolder = localResolveOutputFolder(results, outputFolder)
if nargin >= 2 && ~isempty(outputFolder)
    outputFolder = char(string(outputFolder));
    return;
end

if ~isempty(results) && isfield(results, 'Paths') && isfield(results.Paths, 'Root') && strlength(string(results.Paths.Root)) > 0
    outputFolder = char(string(results.Paths.Root));
else
    outputFolder = fullfile(fileparts(mfilename('fullpath')), 'Generated_Reports');
end
end

function reportData = localBuildReportData(results, sourceInfo, options)
reportData = struct();
reportData.HasResults = ~isempty(results);
reportData.Results = results;
reportData.SourceInfo = sourceInfo;
reportData.Options = options;
reportData.Title = localLocalizedTitle(options.Language);
reportData.Subtitle = localLocalizedSubtitle(options.Language, options.TemplateSubtitle);
reportData.Project = string(options.Project);
reportData.Author = string(options.Author);
reportData.Version = string(options.Version);
reportData.Confidentiality = string(options.Confidentiality);
reportData.Company = string(options.Company);
reportData.DateString = string(options.DateString);
reportData.PlaceholderTag = string(options.PlaceholderTag);
reportData.Language = string(options.Language);
reportData.RunSourceText = localResolveRunSourceText(results, sourceInfo);
reportData.SignalCatalog = localGetNestedTable(results, {'Metadata', 'SignalCatalog'});
reportData.SpecCatalog = localGetNestedTable(results, {'Metadata', 'SpecCatalog'});
reportData.BlockDiagram = localGetNestedTable(results, {'Metadata', 'BlockDiagram'});
reportData.VehicleKPI = localGetTableField(results, 'VehicleKPI');
reportData.SegmentKPI = localGetTableField(results, 'SegmentKPI');
reportData.SegmentSummary = localGetTableField(results, 'SegmentSummary');
reportData.RootCauseRanking = localGetTableField(results, 'RootCauseRanking');
reportData.BadSegmentTable = localGetTableField(results, 'BadSegmentTable');
reportData.OptimizationTable = localGetTableField(results, 'OptimizationTable');
reportData.SignalPresence = localGetTableField(results, 'SignalPresence');
reportData.SpecPresence = localGetTableField(results, 'SpecPresence');
reportData.ExtractionLog = localGetTableField(results, 'ExtractionLog');
reportData.MatInventory = localGetTableField(results, 'MatInventory');
reportData.VehicleNarrative = localGetStringVector(results, 'VehicleNarrative');
reportData.RootCauseNarrative = localGetStringVector(results, 'RootCauseNarrative');
reportData.SubsystemResults = localGetSubsystemResults(results);
reportData.VehicleFigureFiles = localGetVehicleFigureFiles(results);
reportData.VehicleFigureNotes = localGetVehicleFigureNotes(results);
reportData.ThresholdTable = localGetThresholdTable(results);
reportData.Executive = localBuildExecutiveSummary(reportData);
reportData.Abbreviations = localBuildAbbreviationTable();
reportData.SectionMap = localBuildSectionSourceMap();
end

function titleText = localLocalizedTitle(language)
if string(language) == "DE"
    titleText = "eBus Simulations Root-Cause-Analyse";
else
    titleText = "eBus Simulation Root Cause Analysis";
end
end

function subtitleText = localLocalizedSubtitle(language, defaultSubtitle)
if string(language) == "DE"
    subtitleText = "Technische Bewertung auf Fahrzeug- und Subsystemebene";
else
    subtitleText = string(defaultSubtitle);
end
end

function sourceText = localResolveRunSourceText(results, sourceInfo)
sourceParts = strings(0, 1);
sourceParts(end + 1) = "Data source: " + string(sourceInfo.Description);

if strlength(string(sourceInfo.SourcePath)) > 0
    sourceParts(end + 1) = "Loaded from: " + string(sourceInfo.SourcePath);
end

if ~isempty(results) && isfield(results, 'AnalysisData') && isfield(results.AnalysisData, 'RawData') && isfield(results.AnalysisData.RawData, 'MatFilePath')
    sourceParts(end + 1) = "Simulation MAT file: " + string(results.AnalysisData.RawData.MatFilePath);
end
if ~isempty(results) && isfield(results, 'Metadata') && isfield(results.Metadata, 'ExcelFile')
    sourceParts(end + 1) = "Workbook metadata: " + string(results.Metadata.ExcelFile);
end

sourceText = strjoin(sourceParts, '  ');
end

function value = localGetTableField(results, fieldName)
if isempty(results) || ~isfield(results, fieldName) || ~istable(results.(fieldName))
    value = table();
else
    value = results.(fieldName);
end
end

function value = localGetNestedTable(results, pathParts)
value = table();
try
    node = results;
    for iPart = 1:numel(pathParts)
        if ~isfield(node, pathParts{iPart})
            return;
        end
        node = node.(pathParts{iPart});
    end
    if istable(node)
        value = node;
    end
catch
end
end

function value = localGetStringVector(results, fieldName)
if isempty(results) || ~isfield(results, fieldName)
    value = strings(0, 1);
else
    raw = results.(fieldName);
    if isstring(raw)
        value = raw(:);
    elseif iscellstr(raw) || iscell(raw)
        value = string(raw(:));
    else
        value = string(raw);
        value = value(:);
    end
end
end

function subsystemResults = localGetSubsystemResults(results)
if isempty(results) || ~isfield(results, 'SubsystemResults') || isempty(results.SubsystemResults)
    subsystemResults = struct('Name', "", 'Available', false, 'RequiredSignals', {{}}, ...
        'OptionalSignals', {{}}, 'KPITable', table(), 'FigureFiles', strings(0, 1), ...
        'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), 'Suggestions', table());
    subsystemResults(1) = [];
else
    subsystemResults = results.SubsystemResults;
end
end

function files = localGetVehicleFigureFiles(results)
files = strings(0, 1);
if ~isempty(results) && isfield(results, 'VehiclePlots') && isfield(results.VehiclePlots, 'Files')
    files = string(results.VehiclePlots.Files(:));
end
end

function notes = localGetVehicleFigureNotes(results)
notes = strings(0, 1);
if ~isempty(results) && isfield(results, 'VehiclePlots') && isfield(results.VehiclePlots, 'Notes')
    notes = string(results.VehiclePlots.Notes(:));
end
end

function thresholdTable = localGetThresholdTable(results)
thresholdTable = table();
if ~isempty(results) && isfield(results, 'Config') && isfield(results.Config, 'ThresholdTable')
    thresholdTable = results.Config.ThresholdTable;
else
    try
        thresholdTable = RCA_Config().ThresholdTable;
    catch
    end
end
end

function executive = localBuildExecutiveSummary(reportData)
executive = struct();
executive.Why = "This analysis was performed to convert simulation outputs into vehicle-level and subsystem-level engineering findings, connect observed behaviour to likely physical or controls-related causes, and prioritize improvement actions that improve efficiency, performance, and usable range.";
executive.Data = localBuildDataSummary(reportData);
executive.Findings = localBuildTopFindings(reportData);
executive.RootCauses = localBuildTopRootCauses(reportData);
executive.Actions = localBuildTopActions(reportData);
executive.Risk = localBuildRiskSummary(reportData);
end

function text = localBuildDataSummary(reportData)
parts = strings(0, 1);
parts(end + 1) = "The report uses electric bus simulation logged data from MAT files, workbook-based signal and specification metadata from eBus_Model_Info.xlsx, and post-processed KPI / plot / root-cause outputs generated by the MATLAB RCA workflow.";

if height(reportData.SignalPresence) > 0
    presentCount = sum(reportData.SignalPresence.Status == "Present");
    missingCount = sum(contains(reportData.SignalPresence.Status, "Missing"));
    parts(end + 1) = sprintf('Signal audit result: %d present signals and %d missing or optional-missing signals were identified during analysis.', presentCount, missingCount);
end

if height(reportData.MatInventory) > 0
    parts(end + 1) = sprintf('The source MAT file inventory contained %d top-level variables or containers.', height(reportData.MatInventory));
end

text = strjoin(parts, ' ');
end

function findings = localBuildTopFindings(reportData)
findings = strings(0, 1);

if height(reportData.BadSegmentTable) > 0
    count = min(5, height(reportData.BadSegmentTable));
    for iRow = 1:count
        row = reportData.BadSegmentTable(iRow, :);
        findings(end + 1) = sprintf('Bad segment %d between %.1f s and %.1f s is primarily characterized as %s, with %s contributing %.1f%%.', ...
            row.SegmentID, row.StartTime_s, row.EndTime_s, row.IssueType, row.PrimaryCause, row.PrimaryContribution_pct);
    end
end

if isempty(findings) && numel(reportData.VehicleNarrative) > 0
    findings = reportData.VehicleNarrative(1:min(5, numel(reportData.VehicleNarrative)));
end

if isempty(findings)
    findings = [ ...
        "[Insert Finding 1: headline vehicle efficiency issue]"; ...
        "[Insert Finding 2: headline performance or tracking issue]"; ...
        "[Insert Finding 3: dominant loss contributor]"; ...
        "[Insert Finding 4: subsystem interaction issue]"; ...
        "[Insert Finding 5: model or data quality concern]"];
end
end

function causes = localBuildTopRootCauses(reportData)
causes = strings(0, 1);
if height(reportData.RootCauseRanking) > 0
    causeNames = unique(reportData.RootCauseRanking.CauseName, 'stable');
    aggregate = zeros(numel(causeNames), 1);
    for iCause = 1:numel(causeNames)
        mask = reportData.RootCauseRanking.CauseName == causeNames(iCause);
        aggregate(iCause) = sum(reportData.RootCauseRanking.Contribution_pct(mask), 'omitnan');
    end
    [~, order] = sort(aggregate, 'descend');
    count = min(5, numel(order));
    for iCause = 1:count
        causes(end + 1) = sprintf('%s is a repeated likely cause across poor segments, with aggregated contribution score %.1f%%.', ...
            causeNames(order(iCause)), aggregate(order(iCause)));
    end
end

if isempty(causes)
    causes = [ ...
        "[Insert Root Cause 1: example slope / route severity effect]"; ...
        "[Insert Root Cause 2: example battery or power limit effect]"; ...
        "[Insert Root Cause 3: example gear / shift logic effect]"; ...
        "[Insert Root Cause 4: example controller tracking effect]"; ...
        "[Insert Root Cause 5: example auxiliary burden or regen miss]"];
end
end

function actions = localBuildTopActions(reportData)
actions = strings(0, 1);
if height(reportData.OptimizationTable) > 0
    count = min(5, height(reportData.OptimizationTable));
    for iRow = 1:count
        actions(end + 1) = sprintf('%s owner: %s', ...
            reportData.OptimizationTable.Subsystem(iRow), reportData.OptimizationTable.Recommendation(iRow));
    end
end

if isempty(actions)
    actions = [ ...
        "[Insert Action 1: immediate calibration or investigation action]"; ...
        "[Insert Action 2: model fidelity improvement]"; ...
        "[Insert Action 3: controls logic improvement]"; ...
        "[Insert Action 4: design or efficiency improvement]"; ...
        "[Insert Action 5: additional logging / test recommendation]"];
end
end

function riskText = localBuildRiskSummary(reportData)
if height(reportData.BadSegmentTable) > 0
    poorCount = height(reportData.BadSegmentTable);
    riskText = sprintf(['The RCA identified %d materially poor segments or events that warrant engineering follow-up. ', ...
        'The opportunity is to convert repeated route-load, control, and subsystem inefficiency patterns into prioritized actions with direct efficiency, performance, and range benefit.'], poorCount);
else
    riskText = ['Overall vehicle risk/opportunity summary placeholder: quantify the severity of the observed issue set, the expected vehicle-level impact, ', ...
        'and the likely improvement potential once the recommended actions are implemented.'];
end
end

function abbreviations = localBuildAbbreviationTable()
rows = { ...
    'RCA', 'Root Cause Analysis'; ...
    'KPI', 'Key Performance Indicator'; ...
    'SoC', 'State of Charge'; ...
    'Wh/km', 'Energy consumption per kilometre'; ...
    'PI', 'Proportional-Integral'; ...
    'HV', 'High Voltage'; ...
    'BMS', 'Battery Management System'; ...
    'eDrive', 'Electric drive system including inverter and traction motor'; ...
    'MAE', 'Mean Absolute Error'; ...
    'RMSE', 'Root Mean Square Error'};
abbreviations = cell2table(rows, 'VariableNames', {'Abbreviation', 'Definition'});
end

function sectionMap = localBuildSectionSourceMap()
rows = { ...
    '3 Executive Summary', 'Vehicle_Detailed_Analysis results struct', 'VehicleNarrative, RootCauseNarrative, RootCauseRanking, OptimizationTable'; ...
    '9 Simulation and Data Overview', 'RCA_ReadSignalCatalog.m / RCA_LoadMatData.m', 'Metadata workbook parsing, MAT inventory, signal presence'; ...
    '10 Analysis Methodology', 'Vehicle_Detailed_Analysis.m and RCA helper functions', 'Segmentation, KPI framework, root-cause scoring logic'; ...
    '11 Vehicle-Level Assessment', 'RCA_ComputeVehicleKPIs.m / RCA_GenerateVehiclePlots.m', 'Vehicle KPI tables and vehicle figures'; ...
    '12 Subsystem-Level RCA', 'Analyze_*.m', 'Per-subsystem KPI tables, narratives, and figure sets'; ...
    '13 Event-Based Deep Dives', 'RCA_ComputeSegmentKPIs.m / RCA_ComputeRootCauseScores.m / Analyze_Driver.m', 'Segment/event summaries and worst-case drill-downs'; ...
    '14 KPI Dashboard Summary', 'VehicleKPI / SegmentKPI / subsystem KPI tables', 'Report dashboard tables'; ...
    '15 Root Cause Summary Table', 'RootCauseRanking / BadSegmentTable / OptimizationTable', 'Prioritized issue table'; ...
    '18 Appendices', 'Signal catalog / thresholds / extraction log / script files', 'Supporting evidence and reproducibility references'};
sectionMap = cell2table(rows, 'VariableNames', {'ReportSection', 'PrimarySource', 'EvidenceUsed'});
end

function localCreateReportDocument(wordApp, reportData, outputPath, mode)
doc = wordApp.Documents.Add;
cleanupDoc = onCleanup(@() localCloseDoc(doc));
selection = wordApp.Selection;
state = struct();
state.Mode = string(mode);
state.UseActualData = mode ~= "template";
state.IsTemplate = mode == "template";
state.IsSample = mode == "sample";
localCurrentReportLanguage(reportData.Options.Language);
cleanupLanguage = onCleanup(@() localCurrentReportLanguage("EN"));

localConfigureDocument(doc);

localWriteCoverPage(selection, reportData, state);
localInsertPageBreak(selection);
localWriteDocumentControlSection(doc, selection, reportData, state);
localInsertPageBreak(selection);
localWriteExecutiveSummarySection(doc, selection, reportData, state);
localInsertPageBreak(selection);
localWriteTocSection(selection);
localInsertPageBreak(selection);
localWriteFigureAndTableLists(selection);
localInsertPageBreak(selection);
localWriteAbbreviationsSection(doc, selection, reportData);
localInsertPageBreak(selection);
localWriteIntroductionSection(doc, selection, reportData, state);
localWriteSimulationOverviewSection(doc, selection, reportData, state);
localWriteMethodologySection(doc, selection, reportData, state);
localWriteVehicleAssessmentSection(doc, selection, reportData, state);
localWriteSubsystemSection(doc, selection, reportData, state);
localWriteEventDeepDiveSection(doc, selection, reportData, state);
localWriteKpiDashboardSection(doc, selection, reportData, state);
localWriteRootCauseSummarySection(doc, selection, reportData, state);
localWriteRecommendationsSection(selection, reportData, state);
localWriteConclusionSection(selection, reportData, state);
localWriteAppendixSection(doc, selection, reportData, state);

localUpdateAllFields(doc);
localSaveDocument(doc, outputPath);
clear cleanupLanguage
clear cleanupDoc
end

function localConfigureDocument(doc)
cm = 28.3464567;
try
    doc.PageSetup.PaperSize = 7;
catch
    try
        doc.PageSetup.PageWidth = 21.0 * cm;
        doc.PageSetup.PageHeight = 29.7 * cm;
    catch
    end
end
doc.PageSetup.TopMargin = 2.0 * cm;
doc.PageSetup.BottomMargin = 2.0 * cm;
doc.PageSetup.LeftMargin = 2.2 * cm;
doc.PageSetup.RightMargin = 2.0 * cm;

try
    doc.Styles.Item('Normal').Font.Name = 'Calibri';
    doc.Styles.Item('Normal').Font.Size = 11;
    doc.Styles.Item('Title').Font.Name = 'Calibri';
    doc.Styles.Item('Subtitle').Font.Name = 'Calibri';
    doc.Styles.Item('Heading 1').Font.Name = 'Calibri';
    doc.Styles.Item('Heading 2').Font.Name = 'Calibri';
    doc.Styles.Item('Heading 3').Font.Name = 'Calibri';
    doc.Styles.Item('Caption').Font.Name = 'Calibri';
catch
end
end

function localWriteCoverPage(selection, reportData, state)
localApplyStyle(selection, 'Title');
selection.TypeText(char(reportData.Title));
selection.TypeParagraph;

localApplyStyle(selection, 'Subtitle');
selection.TypeText(char(reportData.Subtitle));
selection.TypeParagraph;
selection.TypeParagraph;
selection.TypeParagraph;

localApplyStyle(selection, 'Normal');
localTypeBoldLine(selection, 'Project / Program: ', reportData.Project);
localTypeBoldLine(selection, 'Author: ', reportData.Author);
localTypeBoldLine(selection, 'Date: ', reportData.DateString);
localTypeBoldLine(selection, 'Version: ', reportData.Version);
localTypeBoldLine(selection, 'Confidentiality: ', reportData.Confidentiality);
localTypeBoldLine(selection, 'Company / Department: ', reportData.Company);
selection.TypeParagraph;
selection.TypeParagraph;

localApplyStyle(selection, 'Normal');
if state.IsTemplate
    selection.TypeText('Template use note: replace the cover-page fields with program-specific details before issue for review.');
else
    selection.TypeText(char(reportData.RunSourceText));
end
selection.TypeParagraph;
selection.TypeParagraph;

localApplyStyle(selection, 'Normal');
selection.TypeText('Prepared as an internal vehicle simulation and engineering root-cause analysis report intended for technical management, simulation leads, subsystem owners, and calibration / integration teams.');
selection.TypeParagraph;
end

function localWriteDocumentControlSection(doc, selection, reportData, state)
localAddHeading(selection, '2. Document Control', 1);

selection.TypeText('Document purpose: ');
selection.Font.Bold = true;
selection.TypeText('This report consolidates electric bus simulation evidence into a structured root-cause analysis focused on efficiency, performance, operating behaviour, and subsystem contribution.');
selection.Font.Bold = false;
selection.TypeParagraph;

selection.TypeText('Scope: ');
selection.Font.Bold = true;
if state.IsTemplate
    selection.TypeText('Define the simulation case boundaries, relevant subsystems, assumptions, and exclusions for the study.');
else
    selection.TypeText('Vehicle-level and subsystem-level assessment of the available simulation case(s), with emphasis on event-based RCA, bad-segment evidence, and actionable engineering recommendations.');
end
selection.Font.Bold = false;
selection.TypeParagraph;
selection.TypeParagraph;

localAddHeading(selection, '2.1 Version History', 2);
versionHeaders = {'Version', 'Date', 'Author', 'Description of Change'};
versionRows = {
    char(reportData.Version), char(reportData.DateString), char(reportData.Author), 'Initial automated RCA report issue';
    '0.x', '[Insert Date]', '[Insert Author]', 'Draft / working issue placeholder'};
localAddWordTable(doc, selection, 'Version history table', versionHeaders, versionRows);

localAddHeading(selection, '2.2 Review and Approval', 2);
reviewHeaders = {'Role', 'Name', 'Review Date', 'Status / Comment'};
reviewRows = {
    'Technical Manager', '[Insert Reviewer]', '[Insert Date]', '[Insert Status]';
    'Simulation Lead', '[Insert Reviewer]', '[Insert Date]', '[Insert Status]';
    'Subsystem Owner', '[Insert Reviewer]', '[Insert Date]', '[Insert Status]'};
localAddWordTable(doc, selection, 'Reviewer and approver table', reviewHeaders, reviewRows);
end

function localWriteExecutiveSummarySection(doc, selection, reportData, state)
localAddHeading(selection, '3. Executive Summary', 1);
localWriteLabelParagraph(selection, 'Why this analysis was performed', reportData.Executive.Why);
localWriteLabelParagraph(selection, 'What data was used', reportData.Executive.Data);
localWriteStringList(selection, 'Top 5 critical findings', reportData.Executive.Findings);
localWriteStringList(selection, 'Top 5 likely root causes', reportData.Executive.RootCauses);
localWriteStringList(selection, 'Top recommended actions', reportData.Executive.Actions);
localWriteLabelParagraph(selection, 'Overall vehicle risk / opportunity summary', reportData.Executive.Risk);

summaryTable = localBuildExecutiveSnapshotTable(reportData, state);
localAddWordTable(doc, selection, 'Executive snapshot of key RCA indicators', ...
    {'Metric', 'Value', 'Interpretation'}, summaryTable);
end

function rows = localBuildExecutiveSnapshotTable(reportData, state)
rows = {};

if state.UseActualData && height(reportData.VehicleKPI) > 0
    rows = [rows; localFindKpiRow(reportData.VehicleKPI, 'Trip Distance')];
    rows = [rows; localFindKpiRow(reportData.VehicleKPI, 'Trip Energy Intensity')];
    rows = [rows; localFindKpiRow(reportData.VehicleKPI, 'Battery Discharge Energy')];
end
if state.UseActualData && height(reportData.BadSegmentTable) > 0
    rows(end + 1, :) = {'Bad segment count', num2str(height(reportData.BadSegmentTable)), 'Segments requiring RCA follow-up'};
end

if isempty(rows)
    rows = { ...
        'Overall energy intensity', '[Insert KPI]', 'Headline trip efficiency result'; ...
        'Worst performance issue', '[Insert KPI / event]', 'Largest operational or tracking concern'; ...
        'Dominant root cause', '[Insert Root Cause]', 'Repeated highest-impact contributor'};
end
end

function localWriteTocSection(selection)
localAddHeading(selection, '4. Table of Contents', 1);
localInsertField(selection, 'TOC \o "1-3" \h \z \u');
selection.TypeParagraph;
end

function localWriteFigureAndTableLists(selection)
localAddHeading(selection, '5. List of Figures', 1);
localInsertField(selection, sprintf('TOC \\h \\z \\c "%s"', char(localCaptionLabel('Figure'))));
selection.TypeParagraph;

localAddHeading(selection, '6. List of Tables', 1);
localInsertField(selection, sprintf('TOC \\h \\z \\c "%s"', char(localCaptionLabel('Table'))));
selection.TypeParagraph;
end

function localWriteAbbreviationsSection(doc, selection, reportData)
localAddHeading(selection, '7. Abbreviations / Nomenclature', 1);
selection.TypeText('The following abbreviations are used repeatedly throughout the RCA document.');
selection.TypeParagraph;
localAddWordTable(doc, selection, 'Abbreviations and nomenclature', ...
    {'Abbreviation', 'Definition'}, table2cell(reportData.Abbreviations));
end

function localWriteIntroductionSection(doc, selection, reportData, state)
localAddHeading(selection, '8. Introduction', 1);

localAddHeading(selection, '8.1 Background of eBus Simulation Program', 2);
if state.IsTemplate
    selection.TypeText('[Insert background on the eBus program, simulation objective, and why this simulation case matters in the development plan.]');
else
    selection.TypeText('The eBus simulation program is used to assess vehicle-level energy consumption, traction performance, control behaviour, subsystem losses, and route sensitivity before physical test exposure. This report turns one or more logged simulations into an engineering RCA pack that can be reviewed by vehicle, controls, and subsystem teams using a common evidence base.');
end
selection.TypeParagraph;

localAddHeading(selection, '8.2 Objective of the Root Cause Analysis', 2);
selection.TypeText('The objective is to identify the most important efficiency, performance, operation, and energy-related issues, connect them to likely subsystem contributors, separate observation from inference, and recommend next actions with an explicit confidence statement.');
selection.TypeParagraph;

localAddHeading(selection, '8.3 Questions This Report Answers', 2);
localWriteStringList(selection, '', [ ...
    "Where is the vehicle losing energy or failing to track demand?"; ...
    "Which subsystems contribute most strongly to poor efficiency or weak performance?"; ...
    "Under which event classes or route conditions do the issues occur?"; ...
    "How confident is the RCA, given the available signals and assumptions?"; ...
    "What should be done next in controls, modelling, design, and logging?"]);

localAddHeading(selection, '8.4 Intended Audience', 2);
localWriteStringList(selection, '', [ ...
    "Technical manager"; ...
    "Simulation lead"; ...
    "Subsystem owners"; ...
    "Calibration / controls / vehicle integration engineers"]);

localAddHeading(selection, '8.5 Report Boundaries and Assumptions', 2);
localWriteStringList(selection, '', [ ...
    "The report is based on simulation outputs and available workbook metadata, not direct test measurement."; ...
    "Missing signals do not stop the workflow; instead, affected conclusions are flagged with a limitation note."; ...
    "Thresholds used in segmentation and RCA scoring are heuristics unless tied directly to workbook specifications."; ...
    "Root cause statements are evidence-based hypotheses unless direct causal proof exists in the logged signal set."]);

signalTable = localBuildSignalAvailabilitySnapshot(reportData, state);
localAddWordTable(doc, selection, 'Signal availability snapshot for report context', ...
    {'Status', 'Count', 'Interpretation'}, signalTable);
end

function rows = localBuildSignalAvailabilitySnapshot(reportData, state)
rows = {};
if state.UseActualData && height(reportData.SignalPresence) > 0
    statusList = unique(reportData.SignalPresence.Status, 'stable');
    for iStatus = 1:numel(statusList)
        mask = reportData.SignalPresence.Status == statusList(iStatus);
        rows(end + 1, :) = {char(statusList(iStatus)), num2str(sum(mask)), 'Signal audit count from RCA presence check'}; %#ok<AGROW>
    end
end
if isempty(rows)
    rows = { ...
        'Present', '[Insert Count]', 'Signals available for direct analysis'; ...
        'Missing', '[Insert Count]', 'Signals missing but not fatal to workflow'; ...
        'Optional Missing', '[Insert Count]', 'Signals absent; some RCA detail may be partial'};
end
end

function localWriteSimulationOverviewSection(doc, selection, reportData, state)
localAddHeading(selection, '9. Simulation and Data Overview', 1);

localAddHeading(selection, '9.1 Model Overview', 2);
selection.TypeText('The report uses the Excel workbook as the primary source of model and signal metadata, including subsystem names, signal descriptions, units, block-diagram context, and evaluation expressions. The MAT file is treated as the logged numerical evidence set.');
selection.TypeParagraph;

localAddHeading(selection, '9.2 Simulation Cases / Drive Cycles / Scenarios Analyzed', 2);
if state.UseActualData
    selection.TypeText(char(reportData.RunSourceText));
else
    selection.TypeText('[Insert simulation case description, route name, drive cycle, loading condition, ambient condition, and any variant identifiers.]');
end
selection.TypeParagraph;

localAddHeading(selection, '9.3 Data Sources', 2);
localWriteStringList(selection, '', [ ...
    "Simulation MAT file(s) containing logged bus signals"; ...
    "eBus_Model_Info.xlsx for signal, unit, subsystem, and specification metadata"; ...
    "RCA-generated KPI tables, plots, and root-cause ranking outputs"; ...
    "Vehicle configuration parameters when available from the run or workbook"]);

localAddHeading(selection, '9.4 Logging Overview', 2);
if state.UseActualData && height(reportData.MatInventory) > 0
    selection.TypeText(sprintf('The MAT inventory indicates %d top-level items. The RCA workflow flattened nested structures, identified likely signal containers, checked signal presence against the workbook, and aligned all usable numeric traces to a common time base.', ...
        height(reportData.MatInventory)));
else
    selection.TypeText('[Insert logging overview: main signal containers, sampling basis, and any special logging notes.]');
end
selection.TypeParagraph;

localAddHeading(selection, '9.5 Important Assumptions', 2);
localWriteStringList(selection, '', [ ...
    "Signal interpretation follows workbook-defined descriptions and sign conventions."; ...
    "Battery signals are normalized to the RCA sign convention before energy calculations."; ...
    "If exact formulas are not possible because of missing inputs, practical approximations are used and flagged accordingly."; ...
    "The analysis remains valid for partial signal sets, but confidence reduces when evidence is incomplete."]);

localAddHeading(selection, '9.6 Data Quality Checks Performed', 2);
localWriteStringList(selection, '', [ ...
    "Workbook sheet and header detection"; ...
    "Signal expression evaluation and fallback matching"; ...
    "Signal presence / missing-signal summary generation"; ...
    "Time-base sanitization, duplicate-time handling, and alignment checks"; ...
    "Missing-data-tolerant KPI calculation and plot generation"]);

localAddHeading(selection, '9.7 Known Data Limitations', 2);
if state.UseActualData && height(reportData.ExtractionLog) > 0
    limitText = sprintf('The RCA extraction log contains %d entries and should be reviewed for ambiguous matches, failed workbook expressions, or fallback extraction decisions that reduce confidence.', height(reportData.ExtractionLog));
else
    limitText = 'Known limitations placeholder: record missing signals, ambiguous signal names, model exclusions, and assumptions that reduce RCA certainty.';
end
selection.TypeText(limitText);
selection.TypeParagraph;

overviewTable = localBuildOverviewTable(reportData, state);
localAddWordTable(doc, selection, 'Simulation and data overview summary', {'Item', 'Value', 'Engineering note'}, overviewTable);
end

function rows = localBuildOverviewTable(reportData, state)
rows = {};
if state.UseActualData && height(reportData.SignalPresence) > 0
    rows(end + 1, :) = {'Signal presence rows', num2str(height(reportData.SignalPresence)), 'Workbook-referenced signals checked against the MAT file'}; %#ok<AGROW>
end
if state.UseActualData && height(reportData.SpecPresence) > 0
    rows(end + 1, :) = {'Specification rows', num2str(height(reportData.SpecPresence)), 'Workbook-referenced specifications available for limit checks'}; %#ok<AGROW>
end
if state.UseActualData && height(reportData.VehicleKPI) > 0
    rows(end + 1, :) = {'Vehicle KPI rows', num2str(height(reportData.VehicleKPI)), 'Vehicle-level KPI population'}; %#ok<AGROW>
end
if state.UseActualData && height(reportData.SegmentSummary) > 0
    rows(end + 1, :) = {'Segment summary rows', num2str(height(reportData.SegmentSummary)), 'Segment-wise event or route partitions'}; %#ok<AGROW>
end
if isempty(rows)
    rows = { ...
        'Simulation case ID', '[Insert Case]', 'Unique analysis reference'; ...
        'Drive cycle / route', '[Insert Route]', 'Context for duty-cycle severity'; ...
        'Data quality note', '[Insert Note]', 'Any known logging or signal limitation'};
end
end

function localWriteMethodologySection(doc, selection, reportData, ~)
localAddHeading(selection, '10. Analysis Methodology', 1);

subsections = { ...
    '10.1 Overall RCA Workflow', ...
    '10.2 KPI Calculation Approach', ...
    '10.3 Vehicle-Level Analysis Approach', ...
    '10.4 Subsystem-Level Drill-Down Approach', ...
    '10.5 Event-Based Segmentation Approach', ...
    '10.6 Bad Segment Detection Logic', ...
    '10.7 Correlation / Causality Logic', ...
    '10.8 Rules Used to Classify Probable Root Causes', ...
    '10.9 Confidence Ranking of Findings'};

texts = { ...
    'The workflow reads workbook metadata, inspects the MAT file, checks signal presence, aligns signals to a common time base, derives key traces, computes vehicle and segment KPIs, ranks likely root causes, runs subsystem-specific analyses, and finally assembles evidence into a review-ready report.', ...
    'KPIs are computed using base MATLAB only. Numerical integrations use finite-sample-tolerant logic. Each KPI stores its name, value, unit, category, subsystem, signal basis, and limitation note so downstream reporting remains auditable.', ...
    'Vehicle-level assessment combines speed tracking, energy flow, efficiency, range sensitivity, loss breakdown, gear behaviour, and environmental context to explain what the bus did over the full trip.', ...
    'Subsystem drill-downs run only when the relevant signals are available. Each subsystem section reports role, signals used, key KPIs, observed issue patterns, root cause candidates, and recommended improvements.', ...
    'The RCA partitions the trip into meaningful dynamic or route-context segments using speed, acceleration, slope, stop conditions, and other event cues. Some subsystem analyses also create event-specific partitions such as acceleration, braking, or cruise segments.', ...
    'Bad segments are detected using configured heuristics such as poor tracking, high energy intensity, high loss share, or explicit subsystem-specific thresholds. Threshold values are centralized in RCA_Config.m so they can be reviewed rather than hidden inside the logic.', ...
    'The methodology distinguishes direct observation from inference. Repeated physical patterns are ranked using normalized severity or contribution indicators, but conclusions remain hypotheses unless the logged data provides direct causal proof.', ...
    'Probable root causes are assigned using evidence features such as slope severity, battery limit usage, gear instability, controller tracking error, loss share, and regen recovery performance. Ranking is therefore physics-guided rather than purely statistical.', ...
    'Confidence is reported qualitatively based on evidence breadth and signal completeness. Findings supported by multiple independent signals or repeated across segments are treated with higher confidence than those derived from sparse evidence.'};

for iSection = 1:numel(subsections)
    localAddHeading(selection, subsections{iSection}, 2);
    selection.TypeText(texts{iSection});
    selection.TypeParagraph;
end

if height(reportData.ThresholdTable) > 0
    thresholdRows = localTableToCellRows(reportData.ThresholdTable(1:min(10, height(reportData.ThresholdTable)), :));
    localAddWordTable(doc, selection, 'Representative analysis thresholds and rationale', ...
        reportData.ThresholdTable.Properties.VariableNames, thresholdRows);
end
end

function localWriteVehicleAssessmentSection(doc, selection, reportData, state)
localAddHeading(selection, '11. Vehicle-Level Assessment', 1);

localWriteObservationSection(selection, state, ...
    '11.1 Vehicle Speed Tracking', ...
    localComposeVehicleObservation(reportData, 'tracking'), ...
    'Use the tracking figure and relevant speed-demand KPIs to show whether the vehicle delivers the requested drive cycle or route profile.', ...
    'Poor tracking affects schedule adherence, perceived drivability, and can also reveal upstream power limitation or control arbitration issues.', ...
    localComposeVehicleRootCause(reportData, 'tracking'), ...
    localComposeSeverity(reportData, 'tracking'), ...
    'Review the worst tracking events, confirm whether the issue is route-load-driven or control-limited, and trace limiter ownership.');
localAddRelevantFigure(selection, reportData, state, {'Vehicle_Speed_Tracking', 'Driver_Tracking_Overview'}, ...
    'Vehicle speed tracking over drive cycle');
localAddRelevantTable(doc, selection, reportData.VehicleKPI, {'Tracking', 'Performance'}, ...
    'Vehicle speed tracking and performance KPI summary', reportData.Options.MaxSummaryRows);

localWriteObservationSection(selection, state, ...
    '11.2 Energy Consumption', ...
    localComposeVehicleObservation(reportData, 'energy'), ...
    'Show trip battery discharge, recovered energy, auxiliary demand, and net traction requirement using cumulative and instantaneous energy plots.', ...
    'Energy consumption at vehicle level is the most direct range driver and provides the reference against which subsystem penalties are judged.', ...
    localComposeVehicleRootCause(reportData, 'energy'), ...
    localComposeSeverity(reportData, 'energy'), ...
    'Confirm whether the energy burden is dominated by route severity, auxiliaries, driveline losses, or control-related inefficiency.');
localAddRelevantFigure(selection, reportData, state, {'Vehicle_Energy_Overview'}, ...
    'Battery power and energy flow');
localAddRelevantFigure(selection, reportData, state, {'Vehicle_Energy_Flow_Diagram'}, ...
    'Vehicle energy flow diagram');

localWriteObservationSection(selection, state, ...
    '11.3 Efficiency', ...
    localComposeVehicleObservation(reportData, 'efficiency'), ...
    'Use Wh/km, energy split, and segment ranking evidence to identify where efficiency falls away from the trip baseline.', ...
    'Efficiency degradation typically combines route severity, losses, and control behaviour; it should not be explained by one signal alone.', ...
    localComposeVehicleRootCause(reportData, 'efficiency'), ...
    localComposeSeverity(reportData, 'efficiency'), ...
    'Quantify the recurring loss mechanisms and check whether the issue is systematic across the trip or limited to certain events.');
localAddRelevantFigure(selection, reportData, state, {'Vehicle_Segment_Ranking', 'Vehicle_RootCause_Pareto'}, ...
    'Segment efficiency ranking and recurring RCA contribution');

localWriteObservationSection(selection, state, ...
    '11.4 Range Impact', ...
    'Translate the observed trip behaviour into range sensitivity using trip energy intensity, SoC usage, and identified burden sources such as auxiliary demand or severe route segments.', ...
    'Use range-related KPI or estimated opportunity statements to explain which contributors materially reduce usable distance.', ...
    'Range impact is the management-level translation of efficiency and energy results.', ...
    localComposeVehicleRootCause(reportData, 'range'), ...
    localComposeSeverity(reportData, 'range'), ...
    'Prioritize actions that remove high-frequency or high-energy penalties first because they offer the largest direct range return.');

localWriteObservationSection(selection, state, ...
    '11.5 Performance Limitations', ...
    'Assess whether vehicle performance shortfall appears during acceleration, hill climb, high-speed operation, or low-SoC conditions.', ...
    'Use tracking error, torque limit usage, and route context to distinguish true performance limitation from unrealistic demand or route severity.', ...
    'Performance limitations may point to battery power limit, controller demand limitation, gear selection, or driveline capability gaps.', ...
    localComposeVehicleRootCause(reportData, 'performance'), ...
    localComposeSeverity(reportData, 'performance'), ...
    'Review representative underperformance events using worst-segment dashboards and subsystem drill-down evidence.');

localWriteObservationSection(selection, state, ...
    '11.6 Operational Anomalies', ...
    'Summarize recurring behaviours such as excessive shifting, pedal overlap, instability, or non-physical oscillation.', ...
    'Use event-level and subsystem evidence to show when the behaviour occurs and whether it is likely a model, controls, or integration issue.', ...
    'Operational anomalies often degrade both efficiency and stakeholder confidence in the model quality.', ...
    localComposeVehicleRootCause(reportData, 'operation'), ...
    localComposeSeverity(reportData, 'operation'), ...
    'Separate model-quality issues from genuine control or hardware limitation so the corrective owner is assigned correctly.');

localWriteObservationSection(selection, state, ...
    '11.7 Thermal or Environmental Impact', ...
    'Connect route grade, ambient conditions, and duty-cycle severity to the observed vehicle response.', ...
    'Use the environment subsystem evidence to explain whether energy or performance penalties are route-driven rather than subsystem-driven.', ...
    'Environmental context prevents incorrect blame assignment to propulsion or battery subsystems when the route itself is severe.', ...
    localComposeVehicleRootCause(reportData, 'environment'), ...
    localComposeSeverity(reportData, 'environment'), ...
    'Normalize cross-case comparisons for route severity before drawing design or calibration conclusions.');
localAddRelevantFigure(selection, reportData, state, {'Environment_Overview', 'Environment_Severity_Map'}, ...
    'Environmental severity and route context');

localWriteObservationSection(selection, state, ...
    '11.8 Regeneration Behavior', ...
    'Assess how effectively braking opportunity is converted into recovered electrical energy and whether friction braking dominates recoverable events.', ...
    'Use regen-recovery or brake-energy evidence where available, and explicitly flag any missing signals that reduce confidence.', ...
    'Poor regeneration directly affects energy consumption and may also indicate coordination issues between brake blending, battery acceptance, and control logic.', ...
    localComposeVehicleRootCause(reportData, 'regen'), ...
    localComposeSeverity(reportData, 'regen'), ...
    'Check representative braking events for battery charge-limit interaction, friction-brake dominance, and driver demand arbitration.');

localWriteObservationSection(selection, state, ...
    '11.9 Auxiliary Load Influence', ...
    'Quantify auxiliary demand magnitude, timing, and its share of battery discharge energy.', ...
    'Use auxiliary power traces and energy share metrics to show when non-traction loads dominate the electrical budget.', ...
    'Auxiliary load can materially reduce range even when the propulsion system itself is healthy.', ...
    localComposeVehicleRootCause(reportData, 'auxiliary'), ...
    localComposeSeverity(reportData, 'auxiliary'), ...
    'Review control strategy, duty-cycle dependence, and whether auxiliary peaks coincide with already severe route segments.');

localWriteObservationSection(selection, state, ...
    '11.10 Drive Cycle Sensitivity', ...
    'Summarize how different route states or event classes change the vehicle response and loss mix.', ...
    'Use the segment and event evidence to compare good and bad cases rather than relying on trip-average values only.', ...
    'Drive-cycle sensitivity is essential for deciding whether a fix should target core hardware capability, calibration, or scenario-specific logic.', ...
    localComposeVehicleRootCause(reportData, 'drivecycle'), ...
    localComposeSeverity(reportData, 'drivecycle'), ...
    'Preserve the event-class context when passing actions to subsystem owners, otherwise the requested fix may not target the right operating region.');
end

function localWriteSubsystemSection(doc, selection, reportData, state)
localAddHeading(selection, '12. Subsystem-Level Root Cause Analysis', 1);

selectedSubsystems = localSelectSubsystemsForReport(reportData);
if isempty(selectedSubsystems)
    localAddHeading(selection, '12.1 Subsystem RCA Placeholder', 2);
    selection.TypeText('[Insert subsystem-level RCA findings, including KPI tables, engineering interpretation, and supporting plots.]');
    selection.TypeParagraph;
    return;
end

for iSub = 1:numel(selectedSubsystems)
    sub = selectedSubsystems(iSub);
    headingText = sprintf('12.%d %s Subsystem Root Cause Analysis', iSub, localPrettyName(sub.Name));
    localWriteFocusedSubsystemSection(doc, selection, state, ...
        headingText, ...
        localPrettyName(sub.Name), ...
        sub, ...
        localSubsystemRoleText(sub.Name), ...
        localBuildSubsystemFigureTokenList(sub, reportData.Options.MaxSubsystemFigures), ...
        localBuildSubsystemFigureCaptions(sub, reportData.Options.MaxSubsystemFigures));
end
end

function localWriteEventDeepDiveSection(doc, selection, reportData, state)
localAddHeading(selection, '13. Event-Based Deep Dives', 1);

eventInfo = { ...
    '13.1 Acceleration Events', 'Triggered when positive acceleration or propulsion demand exceeds the configured threshold.', 'Compare good and bad acceleration events using speed tracking, demand, gear, and torque-limit context.'; ...
    '13.2 Braking Events', 'Triggered when deceleration or brake pedal demand exceeds the configured threshold.', 'Compare friction-dominated and regeneration-effective braking cases.'; ...
    '13.3 Cruising Events', 'Triggered during low-acceleration steady-speed operation.', 'Use cruise periods to isolate route load and steady-state control bias.'; ...
    '13.4 Hill Climb / Grade Events', 'Triggered when road slope exceeds uphill or downhill thresholds.', 'Compare grade response with vehicle speed shortfall, gear usage, and limit behaviour.'; ...
    '13.5 High Auxiliary Load Events', 'Triggered when auxiliary power exceeds the configured burden threshold.', 'Assess whether auxiliary load worsens already severe vehicle conditions.'; ...
    '13.6 Low SoC / Power-Limited Events', 'Triggered when SoC is low or battery limits are active.', 'Differentiate true battery-limited response from control or route artefacts.'};

for iEvent = 1:size(eventInfo, 1)
    localAddHeading(selection, eventInfo{iEvent, 1}, 2);
    localWriteLabelParagraph(selection, 'Trigger definition', eventInfo{iEvent, 2});
    localWriteLabelParagraph(selection, 'Representative bad cases', localComposeBadCaseText(reportData, iEvent));
    localWriteLabelParagraph(selection, 'Best cases for comparison', 'Select one or more low-severity events in the same class to show what acceptable behaviour looks like under comparable demand.');
    localWriteLabelParagraph(selection, 'What differentiates them', eventInfo{iEvent, 3});
    localWriteLabelParagraph(selection, 'Root cause reasoning', 'Use the combination of route context, vehicle response, subsystem-specific evidence, and limiter ownership to explain why the poor event is different from the good event.');
end

localAddRelevantFigure(selection, reportData, state, {'Driver_Event_Highlights', 'Driver_Bad_Segments'}, ...
    'Event-based acceleration analysis');
localAddRelevantFigure(selection, reportData, state, {'WorstSegment_', 'Driver_Worst_Segment_01'}, ...
    'Representative worst-case segment dashboard');

deepDiveTable = localBuildDeepDiveTable(reportData);
localAddWordTable(doc, selection, 'Event-based deep-dive shortlist', ...
    {'Event ID', 'Event Type', 'When it occurs', 'Key difference', 'Likely RCA driver'}, deepDiveTable);
end

function localWriteFocusedSubsystemSection(doc, selection, state, headingText, subsystemLabel, sub, roleText, figureTokens, figureCaptions)
localAddHeading(selection, headingText, 2);
localWriteLabelParagraph(selection, 'Role in vehicle behavior', roleText);

if isempty(sub)
    localWriteLabelParagraph(selection, 'Signals used', '[Insert signal list]');
    localWriteLabelParagraph(selection, 'Engineering interpretation', sprintf('[Insert %s subsystem interpretation based on KPI trends, event behaviour, and route context.]', subsystemLabel));
    localWriteLabelParagraph(selection, 'Observed issue patterns', sprintf('[Insert %s issue patterns and RCA comments.]', subsystemLabel));
    localWriteLabelParagraph(selection, 'Recommended next actions', sprintf('[Insert %s follow-up actions.]', subsystemLabel));
    localAddFigure(selection, state, '', sprintf('%s subsystem figure placeholder', subsystemLabel));
    return;
end

localWriteLabelParagraph(selection, 'Signals used', localFormatSignalList(sub.RequiredSignals, sub.OptionalSignals));

if istable(sub.KPITable) && height(sub.KPITable) > 0
    localAddWordTable(doc, selection, sprintf('%s KPI summary', subsystemLabel), ...
        sub.KPITable.Properties.VariableNames, ...
        localTableToCellRows(sub.KPITable(1:min(15, height(sub.KPITable)), :)));
else
    localWriteLabelParagraph(selection, 'Key KPIs', sprintf('[Insert %s KPI summary table.]', subsystemLabel));
end

localWriteLabelParagraph(selection, 'Engineering interpretation', localJoinTextOrPlaceholder(sub.SummaryText, ...
    sprintf('[Insert %s subsystem interpretation based on KPI trends, event behaviour, and route context.]', subsystemLabel)));
localWriteLabelParagraph(selection, 'Observed issue patterns', localComposeSubsystemContributors(sub));
localWriteLabelParagraph(selection, 'Root cause candidates', localComposeSubsystemRootCause(sub));
localWriteLabelParagraph(selection, 'Recommended modeling / logic / calibration improvements', localComposeSubsystemRecommendations(sub));

if numel(string(sub.Warnings)) > 0
    localWriteLabelParagraph(selection, 'Limitations', localJoinTextOrPlaceholder(sub.Warnings, 'No special subsystem limitation note captured.'));
end

for iFig = 1:numel(figureTokens)
    filePath = localFindSubsystemFigure(sub, figureTokens{iFig});
    localAddFigure(selection, state, filePath, figureCaptions{iFig});
end
end

function sub = localFindSubsystemResult(reportData, subsystemName)
sub = [];
for iSub = 1:numel(reportData.SubsystemResults)
    normalizedName = upper(regexprep(string(reportData.SubsystemResults(iSub).Name), '[^A-Za-z0-9]', ''));
    if normalizedName == upper(regexprep(string(subsystemName), '[^A-Za-z0-9]', ''))
        sub = reportData.SubsystemResults(iSub);
        return;
    end
end
end

function selectedSubsystems = localSelectSubsystemsForReport(reportData)
selectedSubsystems = reportData.SubsystemResults;
if isempty(selectedSubsystems)
    return;
end

scope = string(reportData.Options.ActiveSubsystemReportScope);
scope = upper(regexprep(scope(:), '[^A-Za-z0-9]', ''));
scope(scope == "") = [];

preferredOrder = [ ...
    "ENVIRONMENT"; ...
    "DRIVER"; ...
    "POWERTRAINCONTROLLER"; ...
    "ELECTRICDRIVE"; ...
    "TRANSMISSION"; ...
    "FINALDRIVE"; ...
    "PNEUMATICBRAKESYSTEM"; ...
    "VEHICLEDYNAMICS"; ...
    "BATTERY"; ...
    "BATTERYMANAGEMENTSYSTEM"; ...
    "AUXILIARYLOAD"];

if ~(any(scope == "ALL") || isempty(scope))
    keepMask = false(numel(selectedSubsystems), 1);
    for iSub = 1:numel(selectedSubsystems)
        normalizedName = upper(regexprep(string(selectedSubsystems(iSub).Name), '[^A-Za-z0-9]', ''));
        keepMask(iSub) = any(scope == normalizedName);
    end
    selectedSubsystems = selectedSubsystems(keepMask);
end

if isempty(selectedSubsystems)
    return;
end

orderIdx = localOrderSubsystems(selectedSubsystems, preferredOrder);
selectedSubsystems = selectedSubsystems(orderIdx);
end

function roleText = localSubsystemRoleText(subsystemName)
normalizedName = upper(regexprep(string(subsystemName), '[^A-Za-z0-9]', ''));
switch normalizedName
    case "ENVIRONMENT"
        roleText = 'The environment subsystem defines the route and operating context seen by the vehicle. Desired speed, road slope, and ambient temperature explain duty-cycle severity, route load severity, and thermal context, and therefore provide essential evidence before assigning blame to downstream propulsion or control subsystems.';
    case "DRIVER"
        roleText = 'The driver subsystem converts desired vehicle behaviour into accelerator and brake requests. Its PI and feedforward behaviour influences speed tracking quality, event response, and the extent to which route severity is converted into demand on the propulsion and brake systems.';
    case "POWERTRAINCONTROLLER"
        roleText = 'The powertrain controller converts pedal demand and operating context into torque requests and torque-limit management for the electric machines. It shapes driveability, recuperation behaviour, and how close the propulsion system operates to its available envelope.';
    case "ELECTRICDRIVE"
        roleText = 'The electric drive subsystem converts commanded torque into motor torque, speed, electrical power flow, and machine losses. It is a primary determinant of propulsion efficiency, regenerative effectiveness, and machine operating-region quality.';
    case "TRANSMISSION"
        roleText = 'The transmission subsystem transfers summed motor torque to the driveline using the selected gear state and ratio. Its shift behaviour, torque transfer quality, and internal losses directly affect performance and energy efficiency.';
    case "FINALDRIVE"
        roleText = 'The final drive converts gearbox torque into net tractive effort at the axle and road interface. It provides the force path that determines launch capability, hill-climb margin, and force-delivery quality at the vehicle level.';
    case "PNEUMATICBRAKESYSTEM"
        roleText = 'The pneumatic brake subsystem provides friction braking force and braking power dissipation. Its interaction with regenerative braking is critical for energy recovery effectiveness, stopping control, and avoidable brake-energy loss.';
    case "VEHICLEDYNAMICS"
        roleText = 'The vehicle dynamics subsystem combines tractive force, braking force, and route loads into wheel force, acceleration, speed, and position. It is the direct evidence layer for whether the full vehicle is behaving physically and meeting demand.';
    case "BATTERY"
        roleText = 'The battery subsystem supplies and absorbs electrical energy for propulsion and regeneration while introducing losses, voltage behaviour, thermal behaviour, and state-of-charge constraints. It is a first-order driver of range and power capability.';
    case "BATTERYMANAGEMENTSYSTEM"
        roleText = 'The battery management system defines allowable charge and discharge limits in current and power. It governs when the energy storage system becomes the limiting element for performance or recuperation.';
    case "AUXILIARYLOAD"
        roleText = 'The auxiliary subsystem draws non-traction electrical power from the HV system. Its duty-cycle and magnitude determine how much usable propulsion energy and range are consumed by support loads instead of vehicle motion.';
    otherwise
        roleText = 'This subsystem contributes to vehicle behaviour through its logged outputs, KPI trends, and interaction with the rest of the propulsion and vehicle system.';
end
end

function figureTokens = localBuildSubsystemFigureTokenList(sub, maxFigures)
figureFiles = localExistingFiles(string(sub.FigureFiles(:)));
if isempty(figureFiles)
    figureTokens = strings(0, 1);
    return;
end
count = min(numel(figureFiles), maxFigures);
figureTokens = figureFiles(1:count);
end

function figureCaptions = localBuildSubsystemFigureCaptions(sub, maxFigures)
figureFiles = localExistingFiles(string(sub.FigureFiles(:)));
if isempty(figureFiles)
    figureCaptions = strings(0, 1);
    return;
end
count = min(numel(figureFiles), maxFigures);
figureCaptions = strings(count, 1);
for iFile = 1:count
    [~, baseName] = fileparts(char(figureFiles(iFile)));
    captionName = strrep(baseName, '_', ' ');
    figureCaptions(iFile) = sprintf('%s RCA - %s', localPrettyName(sub.Name), captionName);
end
end

function filePath = localFindSubsystemFigure(sub, token)
filePath = '';
figureFiles = localExistingFiles(string(sub.FigureFiles(:)));
if isempty(figureFiles)
    return;
end
mask = contains(lower(figureFiles), lower(string(token)));
idx = find(mask, 1, 'first');
if ~isempty(idx)
    filePath = char(figureFiles(idx));
elseif strlength(string(token)) > 0 && isfile(char(string(token)))
    filePath = char(string(token));
end
end

function rows = localBuildDeepDiveTable(reportData)
rows = {};
if height(reportData.BadSegmentTable) > 0
    count = min(6, height(reportData.BadSegmentTable));
    for iRow = 1:count
        row = reportData.BadSegmentTable(iRow, :);
        rows(end + 1, :) = { ...
            sprintf('SEG-%02d', row.SegmentID), ...
            char(row.IssueType), ...
            sprintf('%.1f s to %.1f s', row.StartTime_s, row.EndTime_s), ...
            char(row.Narrative), ...
            char(row.PrimaryCause)}; %#ok<AGROW>
    end
end
if isempty(rows)
    rows = { ...
        'EVT-01', 'Acceleration', '[Insert Time Window]', '[Insert bad-vs-good difference]', '[Insert likely cause]'; ...
        'EVT-02', 'Braking', '[Insert Time Window]', '[Insert bad-vs-good difference]', '[Insert likely cause]'; ...
        'EVT-03', 'Hill Climb', '[Insert Time Window]', '[Insert bad-vs-good difference]', '[Insert likely cause]'};
end
end

function localWriteKpiDashboardSection(doc, selection, reportData, state)
localAddHeading(selection, '14. KPI Dashboard Summary', 1);

dashboardSpecs = { ...
    '14.1 Vehicle KPIs', reportData.VehicleKPI, {'General', 'Operation', 'Range'}; ...
    '14.2 Efficiency KPIs', reportData.VehicleKPI, {'Efficiency'}; ...
    '14.3 Energy Breakdown KPIs', reportData.VehicleKPI, {'Energy', 'Losses'}; ...
    '14.4 Performance KPIs', reportData.VehicleKPI, {'Performance', 'Tracking'}; ...
    '14.5 Tracking / Control KPIs', localCollectSubsystemKPI(reportData, 'DRIVER'), {'Tracking', 'Demand', 'Feedforward', 'RootCause'}; ...
    '14.6 Regeneration KPIs', reportData.SegmentKPI, {'Regen', 'Energy'}; ...
    '14.7 Subsystem Loss Contribution KPIs', localCollectAllSubsystemKPI(reportData), {'Losses', 'Efficiency'}; ...
    '14.8 Data Quality / Analysis Confidence KPIs', localBuildDataQualityKpiTable(reportData, state), {'DataQuality', 'Confidence'}};

for iBlock = 1:size(dashboardSpecs, 1)
    localAddHeading(selection, dashboardSpecs{iBlock, 1}, 2);
    tbl = dashboardSpecs{iBlock, 2};
    categories = dashboardSpecs{iBlock, 3};
    sliced = localSliceKpiTable(tbl, categories, reportData.Options.MaxSummaryRows);
    if height(sliced) > 0
        localAddWordTable(doc, selection, sprintf('%s summary table', dashboardSpecs{iBlock, 1}), ...
            sliced.Properties.VariableNames, localTableToCellRows(sliced));
    else
        selection.TypeText('[Insert KPI dashboard table for this category.]');
        selection.TypeParagraph;
    end
end
end

function localWriteRootCauseSummarySection(doc, selection, reportData, state)
localAddHeading(selection, '15. Root Cause Summary Table', 1);
selection.TypeText('The table below is intended to be the high-value management and owner handoff sheet. Each line should correspond to one issue that warrants action, with clear evidence and ownership.');
selection.TypeParagraph;

rootCauseTable = localBuildRootCauseSummaryTable(reportData, state);
localAddWordTable(doc, selection, 'Root cause summary and action ownership table', ...
    {'Issue ID', 'Symptom', 'Affected metric / KPI', 'When it occurs', 'Vehicle impact', 'Likely root cause', ...
    'Supporting evidence', 'Confidence level', 'Recommended owner', 'Recommended action', 'Priority'}, rootCauseTable);
end

function rows = localBuildRootCauseSummaryTable(reportData, state)
rows = {};

if state.UseActualData && height(reportData.BadSegmentTable) > 0
    count = min(10, height(reportData.BadSegmentTable));
    for iRow = 1:count
        row = reportData.BadSegmentTable(iRow, :);
        primaryCause = localRowFieldText(row, {'PrimaryCause', 'CauseName'}, '[Insert likely root cause]');
        issueType = localRowFieldText(row, {'IssueType'}, '[Insert Symptom]');
        evidenceText = localRowFieldText(row, {'EvidenceSignals', 'Evidence', 'SignalBasis'}, '[Insert supporting evidence]');
        confidenceText = localRowFieldText(row, {'Confidence', 'ConfidenceLevel'}, 'Medium');
        startTimeText = localRowFieldNumber(row, {'StartTime_s'}, NaN);
        endTimeText = localRowFieldNumber(row, {'EndTime_s'}, NaN);
        recommendedOwner = localLookupRecommendationOwner(reportData, primaryCause);
        recommendedAction = localLookupRecommendation(reportData, primaryCause);
        rows(end + 1, :) = { ...
            sprintf('RCA-%02d', iRow), ...
            issueType, ...
            '[Insert KPI / see supporting tables]', ...
            sprintf('%.1f s to %.1f s', startTimeText, endTimeText), ...
            'Vehicle efficiency / performance / drivability penalty', ...
            primaryCause, ...
            evidenceText, ...
            confidenceText, ...
            recommendedOwner, ...
            recommendedAction, ...
            localPriorityFromConfidence(confidenceText)}; %#ok<AGROW>
    end
end

if isempty(rows)
    rows = { ...
        'RCA-01', '[Insert Symptom]', '[Insert KPI]', '[Insert condition / timing]', '[Insert vehicle impact]', ...
        '[Insert likely root cause]', '[Insert supporting evidence]', '[High / Medium / Low]', ...
        '[Insert owner]', '[Insert action]', '[High / Medium / Low]'};
end
end

function localWriteRecommendationsSection(selection, reportData, state)
localAddHeading(selection, '16. Recommendations', 1);

sections = { ...
    '16.1 Immediate Actions', localFilterRecommendations(reportData, state, 1); ...
    '16.2 Medium-Term Model Improvements', localBuildDefaultRecommendationBlock('model'); ...
    '16.3 Controls / Calibration Improvements', localBuildDefaultRecommendationBlock('controls'); ...
    '16.4 Design Improvement Opportunities', localBuildDefaultRecommendationBlock('design'); ...
    '16.5 Additional Simulations / Tests Required', localBuildDefaultRecommendationBlock('test'); ...
    '16.6 Measurement / Validation Recommendations', localBuildDefaultRecommendationBlock('validation')};

for iSection = 1:size(sections, 1)
    localAddHeading(selection, sections{iSection, 1}, 2);
    localWriteStringList(selection, '', sections{iSection, 2});
end
end

function localWriteConclusionSection(selection, reportData, state)
localAddHeading(selection, '17. Conclusion', 1);
if state.UseActualData
    conclusionText = sprintf(['The RCA converted the available simulation evidence into a structured vehicle-level and subsystem-level finding set. ', ...
        '%s %s The next engineering step is to close the highest-priority issues using targeted calibration, model, design, and logging actions, then rerun the same report flow to confirm whether the fixes move the vehicle in the desired direction.'], ...
        localJoinTextOrPlaceholder(reportData.VehicleNarrative(1:min(2, numel(reportData.VehicleNarrative))), ''), ...
        localJoinTextOrPlaceholder(reportData.RootCauseNarrative(1:min(2, numel(reportData.RootCauseNarrative))), ''));
else
    conclusionText = ['Summarize the main engineering findings, the overall vehicle limitations, the biggest opportunity areas, and the next work packages required to de-risk the program. ', ...
        'Keep the conclusion short, evidence-based, and aligned with the issue table and recommendation section.'];
end
selection.TypeText(conclusionText);
selection.TypeParagraph;
end

function localWriteAppendixSection(doc, selection, reportData, state)
localAddHeading(selection, '18. Appendices', 1);

localAddHeading(selection, '18.1 Detailed Signal List', 2);
if state.UseActualData && height(reportData.SignalPresence) > 0
    signalTable = reportData.SignalPresence(1:min(reportData.Options.MaxAppendixRows, height(reportData.SignalPresence)), :);
    localAddWordTable(doc, selection, 'Detailed signal list and presence status', signalTable.Properties.VariableNames, localTableToCellRows(signalTable));
else
    selection.TypeText('[Insert detailed signal list with subsystem, description, unit, and presence status.]');
    selection.TypeParagraph;
end

localAddHeading(selection, '18.2 Full KPI Definitions', 2);
if state.UseActualData && height(reportData.VehicleKPI) > 0
    kpiTable = reportData.VehicleKPI(1:min(reportData.Options.MaxAppendixRows, height(reportData.VehicleKPI)), :);
    localAddWordTable(doc, selection, 'Vehicle KPI definitions and notes', kpiTable.Properties.VariableNames, localTableToCellRows(kpiTable));
else
    selection.TypeText('[Insert full KPI definition table, including formula basis, units, and note field.]');
    selection.TypeParagraph;
end

localAddHeading(selection, '18.3 Data Cleaning Rules', 2);
localWriteStringList(selection, '', [ ...
    "Signal extraction is attempted using workbook evaluation expressions first and fallback matching second."; ...
    "Time vectors are sanitized for monotonicity and duplicate timestamps before interpolation."; ...
    "Missing or non-finite samples are tolerated in KPI calculations wherever practical."; ...
    "Signals are aligned to a common reference time base; limitations are reported when fallback alignment is used."]);

localAddHeading(selection, '18.4 Scenario / Run Descriptions', 2);
selection.TypeText(char(reportData.RunSourceText));
selection.TypeParagraph;

localAddHeading(selection, '18.5 Extra Plots', 2);
extraFigureFiles = localCollectExtraFigures(reportData);
if ~isempty(extraFigureFiles)
    count = min(6, numel(extraFigureFiles));
    for iFigure = 1:count
        localAddFigure(selection, state, char(extraFigureFiles(iFigure)), sprintf('Appendix supporting figure %d', iFigure));
    end
else
    selection.TypeText('[Insert additional supporting plots not carried in the main body.]');
    selection.TypeParagraph;
end

localAddHeading(selection, '18.6 MATLAB Script References', 2);
localAddWordTable(doc, selection, 'MATLAB script reference map', {'Script / File', 'Role in workflow'}, localBuildScriptReferenceRows());

localAddHeading(selection, '18.7 Assumptions and Limitations', 2);
if state.UseActualData && height(reportData.ExtractionLog) > 0
    selection.TypeText(sprintf('Review the extraction log and signal presence tables for detailed limitations. The current run recorded %d extraction-log entries that may include fallback matches or workbook-evaluation issues.', ...
        height(reportData.ExtractionLog)));
else
    selection.TypeText('List the major modelling, signal, and methodology assumptions together with the limitations they impose on RCA confidence.');
end
selection.TypeParagraph;

localAddHeading(selection, '18.8 Mapping of Report Sections to Source Files', 2);
localAddWordTable(doc, selection, 'Mapping of report sections to source files', ...
    reportData.SectionMap.Properties.VariableNames, localTableToCellRows(reportData.SectionMap));

localAddHeading(selection, '18.9 Style Guide for Future Reuse', 2);
localWriteStringList(selection, '', [ ...
    "Keep the report single-column and avoid decorative layout elements."; ...
    "Ensure every major claim cites a KPI, plot, table, or explicit reasoning path."; ...
    "Distinguish observation, inference, root cause hypothesis, and recommendation explicitly."; ...
    "Use numbered headings consistently and keep figure and table captions descriptive."; ...
    "Preserve appendix traceability so future reports remain reproducible and reviewable."; ...
    "When values are missing, retain the section and use explicit placeholders rather than deleting structure."]);
end

function localWriteObservationSection(selection, state, headingText, observationText, evidenceText, interpretationText, rootCauseText, severityText, nextStepText)
localAddHeading(selection, headingText, 2);
localWriteLabelParagraph(selection, 'Observation', observationText);
localWriteLabelParagraph(selection, 'Evidence', evidenceText);
localWriteLabelParagraph(selection, 'Engineering interpretation', interpretationText);
localWriteLabelParagraph(selection, 'Root cause hypothesis', rootCauseText);
localWriteLabelParagraph(selection, 'Severity / confidence', severityText);
localWriteLabelParagraph(selection, 'Recommended next step', nextStepText);

if state.IsTemplate
    selection.TypeText('Authoring note: keep this subsection concise and evidence-backed. Replace generic statements with direct reference to the actual KPI and figure evidence.');
    selection.TypeParagraph;
end
end

function localAddRelevantFigure(selection, reportData, state, fileTokens, captionText)
filePath = localFindFigureFile(reportData, fileTokens);
localAddFigure(selection, state, filePath, captionText);
end

function localAddRelevantTable(doc, selection, tableValue, categories, captionText, maxRows)
sliced = localSliceKpiTable(tableValue, categories, maxRows);
if height(sliced) > 0
    localAddWordTable(doc, selection, captionText, sliced.Properties.VariableNames, localTableToCellRows(sliced));
else
    selection.TypeText('[Insert supporting KPI table for this subsection.]');
    selection.TypeParagraph;
end
end

function filePath = localFindFigureFile(reportData, fileTokens)
allFiles = localCollectExtraFigures(reportData);
filePath = '';
for iToken = 1:numel(fileTokens)
    mask = contains(lower(allFiles), lower(string(fileTokens{iToken})));
    idx = find(mask, 1, 'first');
    if ~isempty(idx)
        filePath = char(allFiles(idx));
        return;
    end
end
end

function files = localCollectExtraFigures(reportData)
files = localExistingFiles(reportData.VehicleFigureFiles);
for iSub = 1:numel(reportData.SubsystemResults)
    try
        files = [files; localExistingFiles(string(reportData.SubsystemResults(iSub).FigureFiles(:)))]; %#ok<AGROW>
    catch
    end
end
files = unique(files, 'stable');
end

function rows = localBuildScriptReferenceRows()
rows = { ...
    'Vehicle_Detailed_Analysis.m', 'Top-level RCA orchestration and output generation'; ...
    'RCA_ReadSignalCatalog.m', 'Workbook parsing and metadata catalog build'; ...
    'RCA_LoadMatData.m', 'MAT loading, flattening, and inventory'; ...
    'RCA_CheckSignalPresence.m', 'Signal extraction and presence-status audit'; ...
    'RCA_AlignSignalStore.m', 'Reference-time selection and signal alignment'; ...
    'RCA_CreateSegments.m', 'Trip segmentation logic'; ...
    'RCA_ComputeVehicleKPIs.m', 'Vehicle KPI calculations'; ...
    'RCA_ComputeSegmentKPIs.m', 'Segment KPI calculations'; ...
    'RCA_ComputeRootCauseScores.m', 'Root-cause ranking logic'; ...
    'Analyze_*.m', 'Subsystem-specific RCA and figure generation'; ...
    'Generate_eBus_RCA_Word_Report.m', 'Word report generation and document templating'};
end

function localAddWordTable(doc, selection, captionText, headers, rows)
if nargin < 5 || isempty(rows)
    rows = {'[Insert]', '[Insert]'};
    headers = {'Field', 'Value'};
end

headers = cellstr(string(headers(:)'));
rows = localEnsureCellMatrix(rows);
colCount = max(numel(headers), size(rows, 2));
if numel(headers) < colCount
    headers(end + 1:colCount) = {''};
end
if size(rows, 2) < colCount
    rows(:, end + 1:colCount) = {''};
end

localAddCaption(selection, 'Table', captionText);
wordTable = doc.Tables.Add(selection.Range, size(rows, 1) + 1, colCount);
wordTable.Style = 'Table Grid';
wordTable.Borders.Enable = 1;
wordTable.Rows.Alignment = 1;

for iCol = 1:colCount
    wordTable.Cell(1, iCol).Range.Text = headers{iCol};
    wordTable.Cell(1, iCol).Range.Bold = true;
end

for iRow = 1:size(rows, 1)
    for iCol = 1:colCount
        wordTable.Cell(iRow + 1, iCol).Range.Text = localCellToWordText(rows{iRow, iCol});
    end
end

try
    selection.SetRange(doc.Range.End - 1, doc.Range.End - 1);
catch
    try
        selection.EndKey(6);
    catch
        selection.MoveDown;
    end
end
selection.TypeParagraph;
selection.TypeParagraph;
end

function localAddFigure(selection, state, filePath, captionText)
if nargin < 3
    filePath = '';
end
if nargin < 4 || strlength(string(captionText)) == 0
    captionText = 'Figure placeholder';
end

if ~state.IsTemplate && strlength(string(filePath)) > 0 && isfile(filePath)
    inlineShape = selection.InlineShapes.AddPicture(filePath);
    try
        if inlineShape.Width > 470
            scale = 470 / inlineShape.Width;
            inlineShape.Width = inlineShape.Width * scale;
            inlineShape.Height = inlineShape.Height * scale;
        end
    catch
    end
    selection.TypeParagraph;
else
    selection.TypeText(char(localTranslateText(['[Insert Figure] ' char(captionText)])));
    selection.TypeParagraph;
end

localAddCaption(selection, 'Figure', captionText);
selection.TypeParagraph;
end

function localAddCaption(selection, labelName, captionText)
captionLabel = localCaptionLabel(labelName);
try
    invoke(selection, 'InsertCaption', captionLabel, ['. ' char(string(localTranslateText(captionText)))]);
    selection.TypeParagraph;
catch
    localApplyStyle(selection, 'Caption');
    selection.TypeText(sprintf('%s. %s', char(captionLabel), char(string(localTranslateText(captionText)))));
    selection.TypeParagraph;
end
end

function localAddHeading(selection, textValue, level)
try
    selection.Collapse(0);
catch
end
selection.TypeParagraph;
switch level
    case 1
        styleName = 'Heading 1';
    case 2
        styleName = 'Heading 2';
    otherwise
        styleName = 'Heading 3';
end
localApplyStyle(selection, styleName);
selection.TypeText(char(string(localTranslateText(textValue))));
selection.TypeParagraph;
localApplyStyle(selection, 'Normal');
end

function localApplyStyle(selection, styleName)
try
    selection.Style = styleName;
catch
    try
        set(selection, 'Style', styleName);
    catch
    end
end
end

function localWriteLabelParagraph(selection, labelText, bodyText)
localApplyStyle(selection, 'Normal');
selection.Font.Bold = true;
selection.TypeText([char(string(localTranslateText(labelText))) ': ']);
selection.Font.Bold = false;
selection.TypeText(char(string(localTranslateText(bodyText))));
selection.TypeParagraph;
end

function localWriteStringList(selection, headingLabel, values)
values = string(values(:));
values(values == "") = [];
if nargin >= 2 && strlength(string(headingLabel)) > 0
    localWriteLabelParagraph(selection, headingLabel, '');
end
if isempty(values)
    selection.TypeText(char(localTranslateText('- [Insert item]')));
    selection.TypeParagraph;
    return;
end
for iValue = 1:numel(values)
    selection.TypeText(['- ' char(string(localTranslateText(values(iValue))))]);
    selection.TypeParagraph;
end
end

function localInsertPageBreak(selection)
selection.InsertBreak(7);
end

function localInsertField(selection, fieldCode)
range = selection.Range;
range.Fields.Add(range, -1, fieldCode);
selection.TypeParagraph;
end

function localUpdateAllFields(doc)
try
    doc.Fields.Update;
catch
end
try
    for iStory = 1:doc.StoryRanges.Count
        doc.StoryRanges.Item(iStory).Fields.Update;
    end
catch
end
end

function localSaveDocument(doc, outputPath)
wdFormatXMLDocument = 12;
try
    doc.SaveAs2(outputPath, wdFormatXMLDocument);
catch
    doc.SaveAs(outputPath, wdFormatXMLDocument);
end
end

function localCloseDoc(doc)
try
    doc.Close(false);
catch
end
end

function localCleanupWord(wordApp)
try
    wordApp.Quit;
catch
end
try
    delete(wordApp);
catch
end
end

function rows = localEnsureCellMatrix(rows)
if istable(rows)
    rows = localTableToCellRows(rows);
elseif isempty(rows)
    rows = cell(0, 0);
elseif ~iscell(rows)
    rows = cellstr(string(rows));
end
if isvector(rows) && ~isempty(rows)
    rows = rows(:)';
end
end

function rows = localTableToCellRows(tableValue)
if isempty(tableValue) || height(tableValue) == 0
    rows = cell(0, width(tableValue));
    return;
end

rows = cell(height(tableValue), width(tableValue));
for iRow = 1:height(tableValue)
    for iCol = 1:width(tableValue)
        rows{iRow, iCol} = localCellToWordText(tableValue{iRow, iCol});
    end
end
end

function textValue = localCellToWordText(value)
if isstring(value)
    textValue = char(strjoin(value(:)', '; '));
elseif ischar(value)
    textValue = value;
elseif isnumeric(value) || islogical(value)
    if isempty(value)
        textValue = '';
    elseif isscalar(value)
        if isnan(value)
            textValue = 'NaN';
        else
            textValue = num2str(value, '%.6g');
        end
    else
        textValue = mat2str(value);
    end
elseif iscell(value)
    textValue = strjoin(cellfun(@localCellToWordText, value, 'UniformOutput', false), '; ');
else
    try
        textValue = char(string(value));
    catch
        textValue = '[Unprintable Value]';
    end
end
end

function rows = localFindKpiRow(kpiTable, kpiName)
rows = {};
if isempty(kpiTable) || height(kpiTable) == 0 || ~ismember('KPIName', kpiTable.Properties.VariableNames)
    return;
end
mask = contains(lower(kpiTable.KPIName), lower(kpiName), 'IgnoreCase', true);
idx = find(mask, 1, 'first');
if isempty(idx)
    return;
end
rows = {char(kpiTable.KPIName(idx)), sprintf('%s %s', num2str(kpiTable.Value(idx), '%.4g'), char(kpiTable.Unit(idx))), char(kpiTable.StatusNote(idx))};
end

function orderIdx = localOrderSubsystems(subsystems, preferredOrder)
names = strings(numel(subsystems), 1);
for i = 1:numel(subsystems)
    names(i) = upper(regexprep(string(subsystems(i).Name), '[^A-Za-z0-9]', ''));
end
score = inf(numel(subsystems), 1);
for i = 1:numel(subsystems)
    idx = find(preferredOrder == names(i), 1, 'first');
    if ~isempty(idx)
        score(i) = idx;
    else
        score(i) = numel(preferredOrder) + i;
    end
end
[~, orderIdx] = sort(score);
end

function pretty = localPrettyName(nameValue)
pretty = regexprep(char(string(nameValue)), '([a-z])([A-Z])', '$1 $2');
pretty = strrep(pretty, '_', ' ');
pretty = strtrim(pretty);
if strlength(string(pretty)) == 0
    pretty = 'Subsystem';
end
end

function textValue = localFormatSignalList(requiredSignals, optionalSignals)
requiredText = localToSignalString(requiredSignals);
optionalText = localToSignalString(optionalSignals);
parts = strings(0, 1);
if strlength(requiredText) > 0
    parts(end + 1) = "Required: " + requiredText;
end
if strlength(optionalText) > 0
    parts(end + 1) = "Optional/context: " + optionalText;
end
if isempty(parts)
    textValue = "[Insert signal list]";
else
    textValue = strjoin(parts, '  ');
end
end

function textValue = localToSignalString(signalValue)
if isempty(signalValue)
    textValue = "";
else
    textValue = strjoin(string(signalValue(:)'), ', ');
end
end

function textValue = localComposeSubsystemContributors(sub)
if isfield(sub, 'SummaryText') && numel(sub.SummaryText) > 0
    textValue = localJoinTextOrPlaceholder(sub.SummaryText(1:min(3, numel(sub.SummaryText))), '[Insert dominant subsystem issue or loss contributors]');
else
    textValue = '[Insert dominant subsystem issue or loss contributors]';
end
end

function textValue = localComposeSubsystemRootCause(sub)
if isfield(sub, 'Warnings') && numel(sub.Warnings) > 0
    textValue = 'Interpret subsystem findings together with the recorded warnings and missing-signal limitations before drawing a hard causal conclusion.';
else
    textValue = 'Use subsystem KPI patterns, supporting figures, and vehicle-level context to state the most likely subsystem-specific root causes.';
end
end

function textValue = localComposeSubsystemRecommendations(sub)
if isfield(sub, 'Suggestions') && istable(sub.Suggestions) && height(sub.Suggestions) > 0
    try
        textValue = strjoin(string(sub.Suggestions.Recommendation(1:min(3, height(sub.Suggestions)))), '  ');
    catch
        textValue = 'Review the subsystem suggestion table for recommended model, control, or calibration actions.';
    end
else
    textValue = 'Insert subsystem-specific modelling, logic, calibration, or design recommendations.';
end
end

function textValue = localComposeVehicleObservation(reportData, topic)
switch lower(topic)
    case 'tracking'
        if height(reportData.VehicleKPI) > 0
            row = localFindFirstMatchingKpi(reportData.VehicleKPI, {'tracking', 'speed'});
            if ~isempty(row)
                textValue = row;
                return;
            end
        end
        textValue = 'Assess whether actual vehicle speed follows the demanded speed trace across the full drive cycle and across representative event classes.';
    case 'energy'
        textValue = 'Review cumulative battery discharge, auxiliary burden, losses, and recovered energy over the route.';
    case 'efficiency'
        textValue = 'Use trip-level and segment-level Wh/km evidence to identify the most expensive operating regions.';
    case 'range'
        textValue = 'Translate observed energy intensity and SoC usage into a range-sensitive engineering summary.';
    case 'performance'
        textValue = 'Identify where demanded vehicle response cannot be delivered, especially under acceleration or grade.';
    case 'operation'
        textValue = 'Check for anomalous, unstable, or inefficient operating behaviour such as excessive shifting or arbitration issues.';
    case 'environment'
        textValue = 'Use route and ambient context to explain whether the trip itself is severe or benign.';
    case 'regen'
        textValue = 'Compare braking opportunity and recovered energy behaviour to see whether regeneration is being fully utilized.';
    case 'auxiliary'
        textValue = 'Quantify how much auxiliary power demand contributes to the overall energy burden.';
    otherwise
        textValue = 'Compare the vehicle behaviour across different route and demand conditions to understand sensitivity.';
end
end

function textValue = localComposeVehicleRootCause(reportData, topic)
if height(reportData.BadSegmentTable) > 0
    textValue = sprintf('Use the bad-segment RCA table to test whether %s-related issues align with repeated primary causes such as %s.', ...
        topic, strjoin(unique(reportData.BadSegmentTable.PrimaryCause(1:min(3, height(reportData.BadSegmentTable)))), ', '));
else
    textValue = 'State the most likely root-cause hypotheses and distinguish them from direct observations.';
end
end

function textValue = localComposeSeverity(reportData, topic)
if height(reportData.BadSegmentTable) > 0
    textValue = sprintf('Severity should be based on vehicle impact and recurrence. Current RCA has %d explicitly poor segments or events requiring review. Confidence depends on signal completeness and consistency of evidence for the %s topic.', ...
        height(reportData.BadSegmentTable), topic);
else
    textValue = 'Insert a severity and confidence statement using the project review scale.';
end
end

function rowText = localFindFirstMatchingKpi(kpiTable, tokens)
rowText = '';
mask = true(height(kpiTable), 1);
for iToken = 1:numel(tokens)
    mask = mask & contains(lower(kpiTable.KPIName), lower(tokens{iToken}), 'IgnoreCase', true);
end
idx = find(mask, 1, 'first');
if ~isempty(idx)
    rowText = sprintf('%s = %s %s. %s', kpiTable.KPIName(idx), num2str(kpiTable.Value(idx), '%.4g'), kpiTable.Unit(idx), kpiTable.StatusNote(idx));
end
end

function textValue = localComposeBadCaseText(reportData, eventIndex)
if height(reportData.BadSegmentTable) >= eventIndex
    row = reportData.BadSegmentTable(eventIndex, :);
    textValue = sprintf('Representative bad case: segment %d from %.1f s to %.1f s. Primary cause = %s. Narrative: %s', ...
        row.SegmentID, row.StartTime_s, row.EndTime_s, row.PrimaryCause, row.Narrative);
else
    textValue = '[Insert representative bad-case example and its distinguishing evidence.]';
end
end

function sliced = localSliceKpiTable(kpiTable, categories, maxRows)
sliced = table();
if isempty(kpiTable) || height(kpiTable) == 0
    return;
end
if nargin < 2 || isempty(categories)
    sliced = kpiTable(1:min(maxRows, height(kpiTable)), :);
    return;
end

mask = false(height(kpiTable), 1);
for iCat = 1:numel(categories)
    mask = mask | strcmpi(string(kpiTable.Category), string(categories{iCat})) | ...
        contains(lower(string(kpiTable.Category)), lower(string(categories{iCat})), 'IgnoreCase', true);
end
sliced = kpiTable(mask, :);
if height(sliced) > maxRows
    sliced = sliced(1:maxRows, :);
end
end

function kpiTable = localCollectSubsystemKPI(reportData, subsystemName)
kpiTable = table();
for iSub = 1:numel(reportData.SubsystemResults)
    if strcmpi(string(reportData.SubsystemResults(iSub).Name), subsystemName)
        if istable(reportData.SubsystemResults(iSub).KPITable)
            kpiTable = reportData.SubsystemResults(iSub).KPITable;
        end
        return;
    end
end
end

function kpiTable = localCollectAllSubsystemKPI(reportData)
kpiTable = table();
for iSub = 1:numel(reportData.SubsystemResults)
    try
        subTable = reportData.SubsystemResults(iSub).KPITable;
        if istable(subTable) && height(subTable) > 0
            if isempty(kpiTable)
                kpiTable = subTable;
            else
                kpiTable = [kpiTable; subTable]; %#ok<AGROW>
            end
        end
    catch
    end
end
end

function kpiTable = localBuildDataQualityKpiTable(reportData, state)
rows = cell(0, 7);
if state.UseActualData && height(reportData.SignalPresence) > 0
    presentCount = sum(reportData.SignalPresence.Status == "Present");
    missingCount = sum(contains(reportData.SignalPresence.Status, "Missing"));
    rows(end + 1, :) = {'Present signal count', presentCount, 'count', 'DataQuality', 'RCA', 'SignalPresence', 'Signals available for analysis'}; %#ok<AGROW>
    rows(end + 1, :) = {'Missing signal count', missingCount, 'count', 'DataQuality', 'RCA', 'SignalPresence', 'Signals missing or optional-missing'}; %#ok<AGROW>
end
if state.UseActualData && height(reportData.ExtractionLog) > 0
    rows(end + 1, :) = {'Extraction log entries', height(reportData.ExtractionLog), 'count', 'Confidence', 'RCA', 'ExtractionLog', 'Entries may indicate fallback use or ambiguity'}; %#ok<AGROW>
end
if isempty(rows)
    rows = { ...
        'Present signal count', NaN, 'count', 'DataQuality', 'RCA', 'SignalPresence', 'Insert value'; ...
        'Missing signal count', NaN, 'count', 'DataQuality', 'RCA', 'SignalPresence', 'Insert value'; ...
        'Confidence note', NaN, '-', 'Confidence', 'RCA', 'Review', 'Insert qualitative confidence summary'};
end
kpiTable = RCA_FinalizeKPITable(rows);
end

function owner = localLookupRecommendationOwner(reportData, primaryCause)
owner = '[Insert Owner]';
if height(reportData.OptimizationTable) == 0
    return;
end
try
    mask = contains(lower(reportData.OptimizationTable.Evidence), lower(string(primaryCause)), 'IgnoreCase', true) | ...
        contains(lower(reportData.OptimizationTable.Recommendation), lower(string(primaryCause)), 'IgnoreCase', true);
    idx = find(mask, 1, 'first');
    if isempty(idx)
        idx = 1;
    end
    owner = char(reportData.OptimizationTable.Subsystem(idx));
catch
end
end

function action = localLookupRecommendation(reportData, primaryCause)
action = '[Insert Recommended Action]';
if height(reportData.OptimizationTable) == 0
    return;
end
try
    mask = contains(lower(reportData.OptimizationTable.Recommendation), lower(string(primaryCause)), 'IgnoreCase', true) | ...
        contains(lower(reportData.OptimizationTable.Evidence), lower(string(primaryCause)), 'IgnoreCase', true);
    idx = find(mask, 1, 'first');
    if isempty(idx)
        idx = 1;
    end
    action = char(reportData.OptimizationTable.Recommendation(idx));
catch
end
end

function textValue = localRowFieldText(rowTable, candidateNames, defaultValue)
textValue = char(string(defaultValue));
for iName = 1:numel(candidateNames)
    fieldName = candidateNames{iName};
    if ismember(fieldName, rowTable.Properties.VariableNames)
        textValue = localCellToWordText(rowTable.(fieldName)(1));
        if strlength(string(textValue)) == 0
            textValue = char(string(defaultValue));
        end
        return;
    end
end
end

function numberValue = localRowFieldNumber(rowTable, candidateNames, defaultValue)
numberValue = defaultValue;
for iName = 1:numel(candidateNames)
    fieldName = candidateNames{iName};
    if ismember(fieldName, rowTable.Properties.VariableNames)
        candidate = rowTable.(fieldName)(1);
        if isnumeric(candidate) || islogical(candidate)
            numberValue = double(candidate);
        else
            parsed = str2double(string(candidate));
            if ~isnan(parsed)
                numberValue = parsed;
            end
        end
        return;
    end
end
end

function priority = localPriorityFromConfidence(confidenceText)
switch lower(string(confidenceText))
    case "high"
        priority = 'High';
    case "medium"
        priority = 'Medium';
    otherwise
        priority = 'Low';
end
end

function values = localFilterRecommendations(reportData, state, ~)
values = strings(0, 1);
if state.UseActualData && height(reportData.OptimizationTable) > 0
    count = min(5, height(reportData.OptimizationTable));
    for iRow = 1:count
        values(end + 1) = reportData.OptimizationTable.Subsystem(iRow) + ": " + reportData.OptimizationTable.Recommendation(iRow); %#ok<AGROW>
    end
end
if isempty(values)
    values = localBuildDefaultRecommendationBlock('immediate');
end
end

function values = localBuildDefaultRecommendationBlock(mode)
switch lower(mode)
    case 'immediate'
        values = [ ...
            "Confirm the worst-case segments and verify whether they are repeatable across comparable simulation cases."; ...
            "Assign each repeated issue to a subsystem owner with a clear evidence package."; ...
            "Check whether any high-severity finding is caused by missing or ambiguous logging before changing calibration."];
    case 'model'
        values = [ ...
            "Improve subsystem model fidelity where non-physical behaviour or unrealistic limits are suspected."; ...
            "Add missing internal signals that would strengthen causality in future RCA runs."; ...
            "Align workbook metadata and sign-convention documentation with the implemented signal definitions."];
    case 'controls'
        values = [ ...
            "Retune control logic where tracking error, unnecessary transients, or poor arbitration are observed."; ...
            "Review limiter ownership so response delay is separated from true hardware capability limitation."; ...
            "Check shift schedule, torque split, and brake blending calibrations where they materially affect energy or performance."];
    case 'design'
        values = [ ...
            "Review route-sensitive design assumptions such as road-load coefficients, drivetrain efficiency region, and auxiliary architecture."; ...
            "Evaluate whether hardware capability or system sizing is consistent with the demanded operating envelope."; ...
            "Use the RCA to prioritize design changes that remove recurring high-energy penalties."];
    case 'test'
        values = [ ...
            "Rerun the analysis with targeted case variants to separate route severity from subsystem weakness."; ...
            "Add comparison cases for low and high ambient, flat and hilly route, and low and high SoC windows."; ...
            "Create follow-up simulations that isolate the suspected root-cause mechanism rather than only repeating the full mission."];
    otherwise
        values = [ ...
            "Request matching test data or higher-fidelity reference data where simulation confidence is limited."; ...
            "Ensure future data sets include the signals required to confirm or reject the current hypotheses."; ...
            "Track issue closure using the same report structure to preserve comparability across releases."];
end
end

function textValue = localJoinTextOrPlaceholder(values, placeholder)
values = string(values(:));
values(values == "") = [];
if isempty(values)
    textValue = string(placeholder);
else
    textValue = strjoin(values, '  ');
end
end

function files = localExistingFiles(fileList)
fileList = string(fileList(:));
mask = arrayfun(@(x) strlength(x) > 0 && isfile(char(x)), fileList);
files = fileList(mask);
end

function localTypeBoldLine(selection, labelText, valueText)
selection.Font.Bold = true;
selection.TypeText(char(string(localTranslateText(labelText))));
selection.Font.Bold = false;
selection.TypeText(char(string(valueText)));
selection.TypeParagraph;
end

function captionLabel = localCaptionLabel(labelName)
if localCurrentReportLanguage() == "DE"
    switch upper(char(string(labelName)))
        case 'FIGURE'
            captionLabel = "Abbildung";
        case 'TABLE'
            captionLabel = "Tabelle";
        otherwise
            captionLabel = string(localTranslateText(labelName));
    end
else
    captionLabel = string(labelName);
end
end

function language = localCurrentReportLanguage(newLanguage)
persistent currentLanguage
if isempty(currentLanguage)
    currentLanguage = "EN";
end
if nargin >= 1 && strlength(string(newLanguage)) > 0
    currentLanguage = localNormalizeLanguage(newLanguage);
end
language = currentLanguage;
end

function translated = localTranslateText(textValue)
translated = string(textValue);
if localCurrentReportLanguage() ~= "DE"
    return;
end

map = localGermanTranslationMap();
for iValue = 1:numel(translated)
    key = char(translated(iValue));
    if isKey(map, key)
        translated(iValue) = string(map(key));
    end
end
end

function map = localGermanTranslationMap()
persistent translationMap
if ~isempty(translationMap)
    map = translationMap;
    return;
end

keys = { ...
    'Figure', 'Table', ...
    'Project / Program: ', 'Author: ', 'Date: ', 'Version: ', 'Confidentiality: ', 'Company / Department: ', ...
    '2. Document Control', '2.1 Version History', '2.2 Review and Approval', ...
    '3. Executive Summary', '4. Table of Contents', '5. List of Figures', '6. List of Tables', ...
    '7. Abbreviations / Nomenclature', '8. Introduction', ...
    '8.1 Background of eBus Simulation Program', '8.2 Objective of the Root Cause Analysis', ...
    '8.3 Questions This Report Answers', '8.4 Intended Audience', '8.5 Report Boundaries and Assumptions', ...
    '9. Simulation and Data Overview', '9.1 Model Overview', '9.2 Simulation Cases / Drive Cycles / Scenarios Analyzed', ...
    '9.3 Data Sources', '9.4 Logging Overview', '9.5 Important Assumptions', '9.6 Data Quality Checks Performed', ...
    '9.7 Known Data Limitations', '10. Analysis Methodology', '10.1 Overall RCA Workflow', ...
    '10.2 KPI Calculation Approach', '10.3 Vehicle-Level Analysis Approach', '10.4 Subsystem-Level Drill-Down Approach', ...
    '10.5 Event-Based Segmentation Approach', '10.6 Bad Segment Detection Logic', ...
    '10.7 Correlation / Causality Logic', '10.8 Rules Used to Classify Probable Root Causes', ...
    '10.9 Confidence Ranking of Findings', '11. Vehicle-Level Assessment', ...
    '11.1 Vehicle Speed Tracking', '11.2 Energy Consumption', '11.3 Efficiency', '11.4 Range Impact', ...
    '11.5 Performance Limitations', '11.6 Operational Anomalies', '11.7 Thermal or Environmental Impact', ...
    '11.8 Regeneration Behavior', '11.9 Auxiliary Load Influence', '11.10 Drive Cycle Sensitivity', ...
    '12. Subsystem-Level Root Cause Analysis', '13. Event-Based Deep Dives', ...
    '13.1 Acceleration Events', '13.2 Braking Events', '13.3 Cruising Events', ...
    '13.4 Hill Climb / Grade Events', '13.5 High Auxiliary Load Events', ...
    '13.6 Low SoC / Power-Limited Events', '14. KPI Dashboard Summary', '15. Root Cause Summary Table', ...
    '16. Recommendations', '16.1 Immediate Actions', '16.2 Medium-Term Model Improvements', ...
    '16.3 Controls / Calibration Improvements', '16.4 Design Improvement Opportunities', ...
    '16.5 Additional Simulations / Tests Required', '16.6 Measurement / Validation Recommendations', ...
    '17. Conclusion', '18. Appendices', '18.1 Detailed Signal List', '18.2 Full KPI Definitions', ...
    '18.3 Data Cleaning Rules', '18.4 Scenario / Run Descriptions', '18.5 Extra Plots', ...
    '18.6 MATLAB Script References', '18.7 Assumptions and Limitations', ...
    '18.8 Mapping of Report Sections to Source Files', '18.9 Style Guide for Future Reuse', ...
    'Why this analysis was performed', 'What data was used', 'Top 5 critical findings', ...
    'Top 5 likely root causes', 'Top recommended actions', 'Overall vehicle risk / opportunity summary', ...
    'Observation', 'Evidence', 'Engineering interpretation', 'Root cause hypothesis', ...
    'Severity / confidence', 'Recommended next step', 'Trigger definition', ...
    'Representative bad cases', 'Best cases for comparison', 'What differentiates them', ...
    'Root cause reasoning', 'Role in vehicle behavior', 'Signals used', 'Key KPIs', ...
    'Observed issue patterns', 'Root cause candidates', 'Recommended modeling / logic / calibration improvements', ...
    'Limitations', 'Recommended next actions', 'Document purpose', 'Scope', '- [Insert item]'};

values = { ...
    'Abbildung', 'Tabelle', ...
    'Projekt / Programm: ', 'Autor: ', 'Datum: ', 'Version: ', 'Vertraulichkeit: ', 'Firma / Abteilung: ', ...
    '2. Dokumentenlenkung', '2.1 Versionshistorie', '2.2 Prüfung und Freigabe', ...
    '3. Management Summary', '4. Inhaltsverzeichnis', '5. Abbildungsverzeichnis', '6. Tabellenverzeichnis', ...
    '7. Abkürzungen / Nomenklatur', '8. Einleitung', ...
    '8.1 Hintergrund des eBus-Simulationsprogramms', '8.2 Ziel der Root-Cause-Analyse', ...
    '8.3 Fragestellungen dieses Berichts', '8.4 Zielgruppe', '8.5 Berichtsgrenzen und Annahmen', ...
    '9. Simulations- und Datenübersicht', '9.1 Modellübersicht', '9.2 Analysierte Simulationsfälle / Fahrzyklen / Szenarien', ...
    '9.3 Datenquellen', '9.4 Logging-Übersicht', '9.5 Wichtige Annahmen', '9.6 Durchgeführte Datenqualitätsprüfungen', ...
    '9.7 Bekannte Datenbeschränkungen', '10. Analysemethodik', '10.1 Gesamter RCA-Workflow', ...
    '10.2 KPI-Berechnungsmethodik', '10.3 Ansatz der Fahrzeuganalyse', '10.4 Ansatz der Subsystem-Vertiefung', ...
    '10.5 Ereignisbasierte Segmentierung', '10.6 Logik zur Erkennung schlechter Segmente', ...
    '10.7 Korrelations- / Kausalitätslogik', '10.8 Regeln zur Einstufung wahrscheinlicher Ursachen', ...
    '10.9 Vertrauensniveau der Ergebnisse', '11. Fahrzeugbewertung', ...
    '11.1 Fahrgeschwindigkeitsnachführung', '11.2 Energieverbrauch', '11.3 Effizienz', '11.4 Reichweiteneinfluss', ...
    '11.5 Leistungsbegrenzungen', '11.6 Betriebsanomalien', '11.7 Thermischer oder umgebungsbedingter Einfluss', ...
    '11.8 Rekuperationsverhalten', '11.9 Einfluss der Nebenverbraucher', '11.10 Sensitivität gegenüber dem Fahrzyklus', ...
    '12. Root-Cause-Analyse auf Subsystemebene', '13. Ereignisbasierte Detailanalysen', ...
    '13.1 Beschleunigungsereignisse', '13.2 Bremsereignisse', '13.3 Konstantfahrt-Ereignisse', ...
    '13.4 Steigungs- / Gefälleereignisse', '13.5 Ereignisse mit hoher Nebenverbraucherlast', ...
    '13.6 Low-SoC- / leistungsbegrenzte Ereignisse', '14. KPI-Dashboard-Zusammenfassung', '15. Root-Cause-Zusammenfassungstabelle', ...
    '16. Empfehlungen', '16.1 Sofortmaßnahmen', '16.2 Mittelfristige Modellverbesserungen', ...
    '16.3 Regelungs- / Kalibrierungsverbesserungen', '16.4 Konstruktive Verbesserungspotenziale', ...
    '16.5 Zusätzliche Simulationen / Tests', '16.6 Mess- / Validierungsempfehlungen', ...
    '17. Fazit', '18. Anhänge', '18.1 Detaillierte Signalliste', '18.2 Vollständige KPI-Definitionen', ...
    '18.3 Regeln zur Datenbereinigung', '18.4 Szenario- / Laufbeschreibungen', '18.5 Zusätzliche Plots', ...
    '18.6 MATLAB-Skriptverweise', '18.7 Annahmen und Grenzen', ...
    '18.8 Zuordnung der Berichtsabschnitte zu Quelldateien', '18.9 Stilrichtlinie zur Wiederverwendung', ...
    'Warum diese Analyse durchgeführt wurde', 'Welche Daten verwendet wurden', 'Top-5-Schlüsselbefunde', ...
    'Top-5-wahrscheinliche Ursachen', 'Top-Empfehlungen', 'Zusammenfassung von Gesamtrisiko / Chancenbild', ...
    'Beobachtung', 'Nachweis', 'Technische Interpretation', 'Ursachenhypothese', ...
    'Schweregrad / Vertrauen', 'Empfohlener nächster Schritt', 'Triggerdefinition', ...
    'Repräsentative schlechte Fälle', 'Beste Vergleichsfälle', 'Was sie unterscheidet', ...
    'Begründung der Ursachenanalyse', 'Rolle im Fahrzeugverhalten', 'Verwendete Signale', 'Wichtige KPIs', ...
    'Beobachtete Musterausprägungen', 'Mögliche Ursachen', 'Empfohlene Modellierungs- / Logik- / Kalibrierungsverbesserungen', ...
    'Einschränkungen', 'Empfohlene nächste Schritte', 'Dokumentzweck', 'Geltungsbereich', '- [Eintrag einfügen]'};

translationMap = containers.Map(keys, values);
map = translationMap;
end
