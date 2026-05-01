function RCA_ShowVehicleReview(resultsInput)
% RCA_ShowVehicleReview  Display complete vehicle RCA summary and figures.
%
% Usage:
%   RCA_ShowVehicleReview(RCA_Results)
%   RCA_ShowVehicleReview([])

if nargin < 1
    resultsInput = [];
end

results = localResolveResults(resultsInput);

fprintf('\n============================================================\n');
fprintf('Complete Vehicle RCA Review\n');
fprintf('============================================================\n');

if isfield(results, 'Paths') && isfield(results.Paths, 'Root')
    fprintf('Results folder     : %s\n', char(string(results.Paths.Root)));
end

if isfield(results, 'ReferenceInfo') && isstruct(results.ReferenceInfo)
    if isfield(results.ReferenceInfo, 'Source')
        fprintf('Reference time     : %s\n', char(string(results.ReferenceInfo.Source)));
    end
    if isfield(results.ReferenceInfo, 'Method')
        fprintf('Alignment method   : %s\n', char(string(results.ReferenceInfo.Method)));
    end
end

fprintf('Vehicle KPI rows   : %d\n', localTableHeight(results, 'VehicleKPI'));
fprintf('Segment rows       : %d\n', localTableHeight(results, 'SegmentSummary'));
fprintf('Root-cause rows    : %d\n', localTableHeight(results, 'RootCauseRanking'));
fprintf('Bad segments       : %d\n', localTableHeight(results, 'BadSegmentTable'));

fprintf('\nVehicle Interpretation:\n');
if isfield(results, 'VehicleNarrative') && numel(string(results.VehicleNarrative)) > 0
    localPrintTextBlock(string(results.VehicleNarrative));
else
    fprintf('  No vehicle narrative text was recorded.\n');
end

fprintf('\nVehicle KPI Table:\n');
if isfield(results, 'VehicleKPI') && istable(results.VehicleKPI) && height(results.VehicleKPI) > 0
    localPrintPlainTable(results.VehicleKPI);
else
    fprintf('  No vehicle KPI rows were recorded.\n');
end

fprintf('\nSegment Summary:\n');
if isfield(results, 'SegmentSummary') && istable(results.SegmentSummary) && height(results.SegmentSummary) > 0
    localPrintPlainTable(localTopRows(results.SegmentSummary, 12));
else
    fprintf('  No segment summary rows were recorded.\n');
end

fprintf('\nWorst Segments:\n');
if isfield(results, 'BadSegmentTable') && istable(results.BadSegmentTable) && height(results.BadSegmentTable) > 0
    localPrintPlainTable(localTopRows(results.BadSegmentTable, 10));
else
    fprintf('  No bad segments were recorded.\n');
end

fprintf('\nRoot Cause Narrative:\n');
if isfield(results, 'RootCauseNarrative') && numel(string(results.RootCauseNarrative)) > 0
    localPrintTextBlock(string(results.RootCauseNarrative));
else
    fprintf('  No root-cause narrative text was recorded.\n');
end

fprintf('\nRoot Cause Ranking:\n');
if isfield(results, 'RootCauseRanking') && istable(results.RootCauseRanking) && height(results.RootCauseRanking) > 0
    localPrintPlainTable(localTopRows(results.RootCauseRanking, 15));
else
    fprintf('  No root-cause ranking rows were recorded.\n');
end

fprintf('\nOptimization Suggestions:\n');
if isfield(results, 'OptimizationTable') && istable(results.OptimizationTable) && height(results.OptimizationTable) > 0
    localPrintPlainTable(results.OptimizationTable);
else
    fprintf('  No optimization suggestions were recorded.\n');
end

figureFiles = localVehicleFigureFiles(results);
fprintf('\nVehicle Figures:\n');
if isempty(figureFiles)
    fprintf('  No saved vehicle figures were found.\n');
else
    for iFile = 1:numel(figureFiles)
        figureTitle = sprintf('Vehicle RCA Figure %d', iFile);
        fprintf('  %d. %s (%s)\n', iFile, ...
            localMatlabHyperlink(sprintf('RCA_OpenReviewFigure(%s, %s);', ...
            localMatlabStringLiteral(figureFiles(iFile)), localMatlabStringLiteral(figureTitle)), ...
            'Open MATLAB figure'), char(localFigureDisplayName(figureFiles(iFile))));
    end
end

localPrintFollowUpLinks();
end

function results = localResolveResults(resultsInput)
results = [];

if isempty(resultsInput)
    try
        if evalin('base', 'exist(''RCA_Results'',''var'')')
            results = evalin('base', 'RCA_Results');
        end
    catch
    end
elseif isstruct(resultsInput)
    results = resultsInput;
elseif ischar(resultsInput) || (isstring(resultsInput) && isscalar(resultsInput))
    loaded = load(char(string(resultsInput)));
    if isfield(loaded, 'results')
        results = loaded.results;
    elseif isfield(loaded, 'RCA_Results')
        results = loaded.RCA_Results;
    end
end

if isempty(results) || ~isstruct(results) || ~isfield(results, 'VehicleKPI')
    error('RCA_ShowVehicleReview:MissingResults', ...
        'Could not resolve a valid RCA results struct. Run Vehicle_Detailed_Analysis first or pass RCA_Results explicitly.');
end
end

function countValue = localTableHeight(results, fieldName)
countValue = 0;
if isfield(results, fieldName) && istable(results.(fieldName))
    countValue = height(results.(fieldName));
end
end

function subset = localTopRows(inputTable, maxRows)
subset = inputTable(1:min(height(inputTable), maxRows), :);
end

function localPrintTextBlock(lines)
lines = string(lines(:));
lines(lines == "") = [];
for iLine = 1:numel(lines)
    fprintf('  - %s\n', char(lines(iLine)));
end
end

function figureFiles = localVehicleFigureFiles(results)
figureFiles = strings(0, 1);
if ~isfield(results, 'VehiclePlots') || ~isstruct(results.VehiclePlots) || ...
        ~isfield(results.VehiclePlots, 'Files')
    return;
end
figureFiles = localExistingFiles(string(results.VehiclePlots.Files(:)));
end

function files = localExistingFiles(fileList)
fileList = string(fileList(:));
mask = arrayfun(@(x) strlength(x) > 0 && isfile(char(x)), fileList);
files = fileList(mask);
end

function displayName = localFigureDisplayName(filePath)
[folderPath, nameOnly, extension] = fileparts(char(string(filePath)));
displayName = string([nameOnly extension]);
figPath = fullfile(folderPath, [nameOnly '.fig']);
if isfile(figPath)
    displayName = string([nameOnly '.fig']);
end
end

function localPrintPlainTable(inputTable)
if ~istable(inputTable) || isempty(inputTable)
    return;
end

headers = inputTable.Properties.VariableNames;
nRows = height(inputTable);
nCols = width(inputTable);
textData = cell(nRows, nCols);
colWidths = zeros(1, nCols);

for iCol = 1:nCols
    colWidths(iCol) = strlength(string(headers{iCol}));
end

for iRow = 1:nRows
    for iCol = 1:nCols
        textValue = localScalarToText(inputTable{iRow, iCol});
        textData{iRow, iCol} = textValue;
        colWidths(iCol) = max(colWidths(iCol), strlength(string(textValue)));
    end
end

for iCol = 1:nCols
    fprintf('%-*s', colWidths(iCol) + 2, headers{iCol});
end
fprintf('\n');

for iCol = 1:nCols
    fprintf('%s', repmat('-', 1, colWidths(iCol)));
    fprintf('  ');
end
fprintf('\n');

for iRow = 1:nRows
    for iCol = 1:nCols
        fprintf('%-*s', colWidths(iCol) + 2, textData{iRow, iCol});
    end
    fprintf('\n');
end
end

function textValue = localScalarToText(value)
if iscell(value)
    if isempty(value)
        textValue = '';
    elseif numel(value) == 1
        textValue = localScalarToText(value{1});
    else
        parts = cellfun(@localScalarToText, value, 'UniformOutput', false);
        textValue = strjoin(parts, '; ');
    end
elseif isstring(value)
    textValue = char(strjoin(value(:)', '; '));
elseif ischar(value)
    textValue = value;
elseif isnumeric(value) || islogical(value)
    if isempty(value)
        textValue = '';
    elseif isscalar(value)
        textValue = num2str(value, '%.6g');
    else
        textValue = mat2str(value);
    end
else
    try
        textValue = char(string(value));
    catch
        textValue = '[value]';
    end
end
end

function localPrintFollowUpLinks()
fprintf('\nTo View More RCA''s [%s]\n', localMatlabHyperlink( ...
    'RCA_ShowMoreActions(RCA_Results, ''RCA'');', 'Open RCA review list'));
fprintf('To View more KPI''s, Plots, Deep Analysis or Generate Report <a href="matlab:feval(openKpiPlotsViewMore)">[view more...]</a>\n');
end

function hyperlinkText = localMatlabHyperlink(commandText, labelText)
hyperlinkText = sprintf('<a href="matlab:%s">%s</a>', commandText, labelText);
end

function literal = localMatlabStringLiteral(value)
literal = ['''' strrep(char(string(value)), '''', '''''') ''''];
end
