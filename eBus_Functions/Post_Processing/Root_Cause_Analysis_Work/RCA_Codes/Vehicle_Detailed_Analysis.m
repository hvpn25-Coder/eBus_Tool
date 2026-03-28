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

fprintf('\n============================================================\n');
fprintf('Electric Bus Root Cause Analysis\n');
fprintf('MAT file : %s\n', matFilePath);
fprintf('Excel    : %s\n', excelFilePath);
fprintf('Output   : %s\n', outputPaths.Root);
fprintf('============================================================\n');

metadata = RCA_ReadSignalCatalog(excelFilePath, config);
rawData = RCA_LoadMatData(matFilePath, config);
[signals, specs, signalPresence, specPresence, extractionLog] = RCA_CheckSignalPresence(metadata, rawData, config);
[signals, referenceInfo] = RCA_AlignSignalStore(signals, rawData, config);
derived = RCA_BuildDerivedSignals(signals, specs, referenceInfo, config);
segments = RCA_CreateSegments(derived, config);
[vehicleKPI, vehicleNarrative] = RCA_ComputeVehicleKPIs(derived, signals, specs, signalPresence, config);
[segmentKPI, segmentSummary] = RCA_ComputeSegmentKPIs(derived, signals, specs, segments, config);
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

vehiclePlots = RCA_GenerateVehiclePlots(analysisData, outputPaths, config);
subsystemResults = localRunSubsystemAnalyses(analysisData, outputPaths, config);

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

function subsystemResults = localRunSubsystemAnalyses(analysisData, outputPaths, config)
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
    try
        subsystemResults(iAnalyzer) = analyzers{iAnalyzer}(analysisData, outputPaths, config);
    catch analysisException
        subsystemResults(iAnalyzer).Name = string(func2str(analyzers{iAnalyzer}));
        subsystemResults(iAnalyzer).Warnings = "Subsystem analysis failed: " + string(analysisException.message);
    end
end
end
