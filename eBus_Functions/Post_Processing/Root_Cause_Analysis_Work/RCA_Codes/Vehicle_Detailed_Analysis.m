function results = Vehicle_Detailed_Analysis(matFilePath, excelFilePath, outputRoot)
% Vehicle_Detailed_Analysis  Electric bus root cause analysis entry point.
%
% README
% 1. Usage:
%    results = Vehicle_Detailed_Analysis;
%    results = Vehicle_Detailed_Analysis('C:\path\simulation.mat');
%    results = Vehicle_Detailed_Analysis(matFilePath, excelFilePath, outputFolder);
%
% 2. Expected MAT input:
%    The MAT file may contain top-level structs, nested structs, numeric
%    arrays, or timeseries-like data. The workflow inspects the MAT content,
%    builds an evaluation context, and tries workbook equations first.
%
% 3. Excel metadata usage:
%    The workbook is treated as the signal/specification source of truth.
%    Signal and specification evaluation columns are parsed defensively and
%    used in priority order. If an evaluation fails, fallback matching is
%    attempted and logged without stopping the workflow.
%
% 4. Outputs:
%    CSV tables, PNG figures, a MAT results file, and a text summary are
%    written to the output folder. The main result struct is also assigned
%    to the MATLAB base workspace as RCA_Results.

config = RCA_Config();
[matFilePath, excelFilePath, outputRoot] = localResolveInputs(matFilePath, excelFilePath, outputRoot, config);
outputPaths = localInitializeOutputFolders(outputRoot, matFilePath, config);
progressState = localCreateProgressBar('Electric Bus RCA', 'Initializing RCA workflow...');
progressCleanup = onCleanup(@() localCloseProgressBar(progressState)); %#ok<NASGU>
totalSteps = 11;

fprintf('\n============================================================\n');
fprintf('Electric Bus Root Cause Analysis\n');
fprintf('MAT file : %s\n', matFilePath);
fprintf('Excel    : %s\n', excelFilePath);
fprintf('Output   : %s\n', outputPaths.Root);
fprintf('============================================================\n');

localUpdateProgressBar(progressState, 1, totalSteps, 'Reading workbook metadata');
metadata = RCA_ReadSignalCatalog(excelFilePath, config);
localUpdateProgressBar(progressState, 2, totalSteps, 'Loading MAT file and inventory');
rawData = RCA_LoadMatData(matFilePath, config);
localUpdateProgressBar(progressState, 3, totalSteps, 'Checking signal and specification presence');
[signals, specs, signalPresence, specPresence, extractionLog] = RCA_CheckSignalPresence(metadata, rawData, config);
localUpdateProgressBar(progressState, 4, totalSteps, 'Aligning signals to reference time');
[signals, referenceInfo] = RCA_AlignSignalStore(signals, rawData, config, metadata.TimeSignalNames);
localUpdateProgressBar(progressState, 5, totalSteps, 'Building derived signals');
derived = RCA_BuildDerivedSignals(signals, specs, referenceInfo, config);
derived = localAttachRunContextMetrics(derived, rawData);
localUpdateProgressBar(progressState, 6, totalSteps, 'Creating trip segments');
segments = RCA_CreateSegments(derived, config);
localUpdateProgressBar(progressState, 7, totalSteps, 'Computing vehicle KPI');
[vehicleKPI, vehicleNarrative] = RCA_ComputeVehicleKPIs(derived, signals, specs, signalPresence, config);
localUpdateProgressBar(progressState, 8, totalSteps, 'Computing segment KPI');
[segmentKPI, segmentSummary] = RCA_ComputeSegmentKPIs(derived, signals, specs, segments, config);
localUpdateProgressBar(progressState, 9, totalSteps, 'Computing root-cause ranking');
[rootCauseRanking, badSegmentTable, rootCauseNarrative, optimizationTable] = ...
    RCA_ComputeRootCauseScores(derived, signals, specs, segmentSummary, config);

analysisData = struct();
analysisData.Config = config;
analysisData.Metadata = metadata;
analysisData.RawData = rawData;
analysisData.Signals = signals;
analysisData.Specs = specs;
analysisData.ReferenceInfo = referenceInfo;
analysisData.Derived = derived;
analysisData.Segments = segments;
analysisData.SignalPresence = signalPresence;
analysisData.SpecPresence = specPresence;
analysisData.ExtractionLog = extractionLog;
analysisData.VehicleKPI = vehicleKPI;
analysisData.SegmentKPI = segmentKPI;
analysisData.SegmentSummary = segmentSummary;
analysisData.RootCauseRanking = rootCauseRanking;
analysisData.BadSegmentTable = badSegmentTable;
analysisData.OptimizationTable = optimizationTable;

localUpdateProgressBar(progressState, 10, totalSteps, 'Generating figures and subsystem RCA');
hiddenFigureCleanup = localDisableInteractiveFigures(); %#ok<NASGU>
vehiclePlots = RCA_GenerateVehiclePlots(analysisData, outputPaths, config);
subsystemResults = localRunSubsystemAnalyses(analysisData, outputPaths, config, progressState, 10, totalSteps);

results = struct();
results.Config = config;
results.Paths = outputPaths;
results.Metadata = metadata;
results.MatInventory = rawData.InventoryTable;
results.ReferenceInfo = referenceInfo;
results.SignalPresence = signalPresence;
results.SpecPresence = specPresence;
results.ExtractionLog = extractionLog;
results.VehicleKPI = vehicleKPI;
results.SegmentKPI = segmentKPI;
results.SegmentSummary = segmentSummary;
results.RootCauseRanking = rootCauseRanking;
results.BadSegmentTable = badSegmentTable;
results.OptimizationTable = optimizationTable;
results.VehicleNarrative = vehicleNarrative;
results.RootCauseNarrative = rootCauseNarrative;
results.VehiclePlots = vehiclePlots;
results.SubsystemResults = subsystemResults;
results.AnalysisData = analysisData;
results.ReportOutput = struct('ReportFile', "", 'TemplateFile', "", 'SampleFile', "", ...
    'OutputFolder', string(outputPaths.Root), 'Source', struct('HasResults', false));

localUpdateProgressBar(progressState, 11, totalSteps, 'Saving RCA outputs');
RCA_SaveOutputs(results, outputPaths, config);

assignin('base', 'RCA_Results', results);
assignin('base', 'RCA_SignalPresence', signalPresence);
assignin('base', 'RCA_VehicleKPI', vehicleKPI);
assignin('base', 'RCA_SegmentSummary', segmentSummary);
assignin('base', 'RCA_RootCauseRanking', rootCauseRanking);

fprintf('\nSignal presence summary:\n');
disp(signalPresence(:, {'SignalName', 'Description', 'Unit', 'Status', 'Subsystem'}));

fprintf('\nVehicle RCA completed. Key outputs:\n');
fprintf('  Vehicle KPI rows     : %d\n', height(vehicleKPI));
fprintf('  Segment summary rows : %d\n', height(segmentSummary));
fprintf('  Root-cause rows      : %d\n', height(rootCauseRanking));
fprintf('  Output folder        : %s\n', outputPaths.Root);
localPrintActionLinks();
localUpdateProgressBar(progressState, totalSteps, totalSteps, 'RCA execution completed');
localCloseProgressBar(progressState);
end

function [matFilePath, excelFilePath, outputRoot] = localResolveInputs(matFilePath, excelFilePath, outputRoot, config)
if nargin < 1 || isempty(matFilePath)
    [fileName, filePath] = uigetfile('*.mat', 'Select electric bus simulation MAT file');
    if isequal(fileName, 0)
        error('Vehicle_Detailed_Analysis:NoMatFile', 'No MAT file was selected.');
    end
    matFilePath = fullfile(filePath, fileName);
end

if nargin < 2 || isempty(excelFilePath)
    excelFilePath = config.General.DefaultExcelFile;
end

if nargin < 3 || isempty(outputRoot)
    outputRoot = fullfile(fileparts(mfilename('fullpath')), config.General.DefaultResultsFolder);
end

if ~isfile(matFilePath)
    error('Vehicle_Detailed_Analysis:MissingMatFile', 'MAT file not found: %s', matFilePath);
end
if ~isfile(excelFilePath)
    error('Vehicle_Detailed_Analysis:MissingExcelFile', 'Excel file not found: %s', excelFilePath);
end
end

function outputPaths = localInitializeOutputFolders(outputRoot, matFilePath, config)
[~, matBaseName] = fileparts(matFilePath);
runFolder = sprintf('%s_RCA_%s', matBaseName, char(string(datetime('now'), config.General.TimestampFormat)));

outputPaths = struct();
outputPaths.Root = fullfile(outputRoot, runFolder);
outputPaths.Tables = fullfile(outputPaths.Root, 'tables');
outputPaths.Logs = fullfile(outputPaths.Root, 'logs');
outputPaths.Figures = fullfile(outputPaths.Root, 'figures');
outputPaths.FiguresVehicle = fullfile(outputPaths.Figures, 'vehicle');
outputPaths.FiguresSubsystem = fullfile(outputPaths.Figures, 'subsystems');

paths = struct2cell(outputPaths);
for iPath = 1:numel(paths)
    if ~exist(paths{iPath}, 'dir')
        mkdir(paths{iPath});
    end
end
end

function derived = localAttachRunContextMetrics(derived, rawData)
derived.targetDistance_km = NaN;
derived.targetDistanceSource = "";
if nargin < 2 || ~isstruct(rawData) || ~isfield(rawData, 'Workspace') || ~isstruct(rawData.Workspace)
    return;
end

[targetDistance, sourceName] = localResolveWorkspaceScalar(rawData.Workspace, ...
    {'cfg_target_distance', 'target_distance', 'targetDistance', 'cfgTargetDistance', 'targetDist'});
if ~isfinite(targetDistance)
    return;
end

targetDistance_km = targetDistance;
actualDistance_km = NaN;
if isfield(derived, 'tripDistance_km')
    actualDistance_km = derived.tripDistance_km;
end

% Configuration distances may be exported in m or km. Convert obvious
% metre-scale values so target-vs-actual KPI remain physically meaningful.
if abs(targetDistance_km) > 1000 && (isnan(actualDistance_km) || actualDistance_km < 1000)
    targetDistance_km = targetDistance_km / 1000;
elseif isfinite(actualDistance_km) && actualDistance_km > 0 && abs(targetDistance_km) > 10 * actualDistance_km
    targetDistance_km = targetDistance_km / 1000;
end

derived.targetDistance_km = targetDistance_km;
derived.targetDistanceSource = sourceName;
end

function [value, sourceName] = localResolveWorkspaceScalar(workspace, candidateNames)
value = NaN;
sourceName = "";
for iName = 1:numel(candidateNames)
    name = char(candidateNames{iName});
    [candidateValue, ok] = localTryEvaluateWorkspaceExpression(workspace, name);
    if ok && isnumeric(candidateValue) && isscalar(candidateValue) && isfinite(double(candidateValue))
        value = double(candidateValue);
        sourceName = string(name);
        return;
    end
end
end

function [value, ok] = localTryEvaluateWorkspaceExpression(workspace, expressionText)
value = [];
ok = false;
if ~isstruct(workspace) || strlength(string(expressionText)) == 0
    return;
end

workspaceFields = fieldnames(workspace);
for iField = 1:numel(workspaceFields)
    fieldName = workspaceFields{iField};
    if isvarname(fieldName)
        eval([fieldName ' = workspace.(fieldName);']); %#ok<EVLDIR>
    end
end

try
    value = eval(char(expressionText)); %#ok<EVLDIR>
    ok = true;
catch
    value = [];
    ok = false;
end
end

function subsystemResults = localRunSubsystemAnalyses(analysisData, outputPaths, config, progressState, progressStep, totalSteps)
analyzers = { ...
    @Analyze_Environment, ...
    @Analyze_Driver, ...
    @Analyze_PowerTrainController, ...
    @Analyze_ElectricDrive, ...
    @Analyze_Transmission, ...
    @Analyze_FinalDrive, ...
    @Analyze_PneumaticBrakeSystem, ...
    @Analyze_VehicleDynamics, ...
    @Analyze_Battery, ...
    @Analyze_BatteryManagementSystem, ...
    @Analyze_AuxiliaryLoad};

subsystemResults = repmat(struct('Name', "", 'Available', false, 'RequiredSignals', {{}}, ...
    'OptionalSignals', {{}}, 'KPITable', RCA_FinalizeKPITable([]), ...
    'FigureFiles', strings(0, 1), 'SummaryText', strings(0, 1), ...
    'Warnings', strings(0, 1), 'Suggestions', RCA_MakeSuggestionTable("", strings(0, 1), strings(0, 1))), ...
    numel(analyzers), 1);

for iAnalyzer = 1:numel(analyzers)
    localUpdateProgressBar(progressState, progressStep - 0.5 + 0.5 * (iAnalyzer / max(numel(analyzers), 1)), totalSteps, ...
        sprintf('Running subsystem RCA (%d/%d): %s', iAnalyzer, numel(analyzers), func2str(analyzers{iAnalyzer})));
    try
        subsystemResults(iAnalyzer) = analyzers{iAnalyzer}(analysisData, outputPaths, config);
        subsystemResults(iAnalyzer) = RCA_ApplySpecificationContext(subsystemResults(iAnalyzer), analysisData);
    catch analysisException
        subsystemResults(iAnalyzer).Name = string(func2str(analyzers{iAnalyzer}));
        subsystemResults(iAnalyzer).Warnings = "Subsystem analysis failed: " + string(analysisException.message);
    end
end
end

function progressState = localCreateProgressBar(titleText, initialMessage)
progressState = struct('Handle', [], 'Enabled', false);
try
    if usejava('desktop') && feature('ShowFigureWindows')
        progressState.Handle = waitbar(0, initialMessage, 'Name', titleText, ...
            'CreateCancelBtn', '', 'WindowStyle', 'normal');
        progressState.Enabled = ishghandle(progressState.Handle);
    end
catch
    progressState = struct('Handle', [], 'Enabled', false);
end
end

function localUpdateProgressBar(progressState, currentStep, totalSteps, messageText)
if isempty(progressState) || ~isstruct(progressState) || ~isfield(progressState, 'Enabled') || ~progressState.Enabled
    return;
end
try
    fraction = max(0, min(1, double(currentStep) / max(double(totalSteps), 1)));
    waitbar(fraction, progressState.Handle, messageText);
    drawnow limitrate;
catch
end
end

function localCloseProgressBar(progressState)
if isempty(progressState) || ~isstruct(progressState) || ~isfield(progressState, 'Enabled') || ~progressState.Enabled
    return;
end
try
    if isgraphics(progressState.Handle)
        delete(progressState.Handle);
        drawnow;
    end
catch
end
end

function cleanupHandle = localDisableInteractiveFigures()
cleanupHandle = [];
try
    originalVisible = get(groot, 'defaultFigureVisible');
    set(groot, 'defaultFigureVisible', 'off');
    cleanupHandle = onCleanup(@() set(groot, 'defaultFigureVisible', originalVisible));
catch
    cleanupHandle = [];
end
end

function localPrintActionLinks()
fprintf('\nInteractive Actions:\n');
fprintf('  %s\n', localMatlabHyperlink( ...
    'Generate_eBus_RCA_Word_Report(RCA_Results, RCA_Results.Paths.Root, struct(''Language'',''EN''));', ...
    'Generate Word report (English)'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'Generate_eBus_RCA_Word_Report(RCA_Results, RCA_Results.Paths.Root, struct(''Language'',''DE''));', ...
    'Generate Word report (German)'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowVehicleReview(RCA_Results);', ...
    'Open Complete Vehicle RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''DRIVER'');', ...
    'Open Driver RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''ENVIRONMENT'');', ...
    'Open Environment RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''POWER TRAIN CONTROLLER'');', ...
    'Open Powertrain Controller RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''ELECTRIC DRIVE'');', ...
    'Open Electric Drive RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''TRANSMISSION'');', ...
    'Open Transmission RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''FINAL DRIVE'');', ...
    'Open Final Drive RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''PNEUMATIC BRAKE SYSTEM'');', ...
    'Open Pneumatic Brake RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''VEHICLE DYNAMICS'');', ...
    'Open Vehicle Dynamics RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''BATTERY'');', ...
    'Open Battery RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''BATTERY MANAGEMENT SYSTEM'');', ...
    'Open Battery Management System RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''AUXILIARY LOAD'');', ...
    'Open Auxiliary RCA review'));
end

function hyperlinkText = localMatlabHyperlink(commandText, labelText)
hyperlinkText = sprintf('<a href="matlab:%s">%s</a>', commandText, labelText);
end
