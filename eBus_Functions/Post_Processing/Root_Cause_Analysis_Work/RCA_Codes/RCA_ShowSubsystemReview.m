function RCA_ShowSubsystemReview(resultsInput, subsystemName)
% RCA_ShowSubsystemReview  Display one subsystem RCA summary and figures.
%
% Usage:
%   RCA_ShowSubsystemReview(RCA_Results, 'DRIVER')
%   RCA_ShowSubsystemReview(RCA_Results, 'ENVIRONMENT')
%   RCA_ShowSubsystemReview([], 'DRIVER')

if nargin < 1
    resultsInput = [];
end
if nargin < 2 || isempty(subsystemName)
    error('RCA_ShowSubsystemReview:MissingSubsystem', ...
        'Provide a subsystem name such as ''DRIVER'' or ''ENVIRONMENT''.');
end

results = localResolveResults(resultsInput);
subsystemKey = upper(regexprep(string(subsystemName), '[^A-Za-z0-9]', ''));
sub = localFindSubsystem(results, subsystemKey);

if isempty(sub)
    availableNames = localAvailableSubsystemNames(results);
    error('RCA_ShowSubsystemReview:SubsystemNotFound', ...
        'Subsystem %s was not found in the RCA results. Available subsystem names: %s', ...
        char(string(subsystemName)), char(strjoin(availableNames, ', ')));
end

fprintf('\n============================================================\n');
fprintf('%s RCA Review\n', localPrettyName(sub.Name));
fprintf('============================================================\n');
fprintf('Available          : %s\n', string(sub.Available));
fprintf('Required signals   : %s\n', char(localSignalList(sub.RequiredSignals)));
fprintf('Optional/context   : %s\n', char(localSignalList(sub.OptionalSignals)));

if isfield(results, 'Paths') && isfield(results.Paths, 'Root')
    fprintf('Results folder     : %s\n', char(string(results.Paths.Root)));
end

fprintf('\nEngineering Interpretation:\n');
if numel(string(sub.SummaryText)) > 0
    localPrintTextBlock(string(sub.SummaryText));
else
    fprintf('  No subsystem summary text was recorded.\n');
end

if numel(string(sub.Warnings)) > 0
    fprintf('\nWarnings / Limitations:\n');
    localPrintTextBlock(string(sub.Warnings));
end

fprintf('\nKPI Table:\n');
if istable(sub.KPITable) && height(sub.KPITable) > 0
    localPrintPlainTable(sub.KPITable);
else
    fprintf('  No KPI rows were recorded for this subsystem.\n');
end

if isfield(sub, 'Suggestions') && istable(sub.Suggestions) && height(sub.Suggestions) > 0
    fprintf('\nSuggestions:\n');
    localPrintPlainTable(sub.Suggestions);
end

figureFiles = localExistingFiles(string(sub.FigureFiles(:)));
fprintf('\nFigures:\n');
if isempty(figureFiles)
    fprintf('  No saved figures were found for this subsystem.\n');
else
    for iFile = 1:numel(figureFiles)
        fprintf('  %d. %s\n', iFile, char(figureFiles(iFile)));
        localOpenImageFigure(figureFiles(iFile), sprintf('%s RCA Figure %d', localPrettyName(sub.Name), iFile));
    end
end
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

if isempty(results) || ~isstruct(results) || ~isfield(results, 'SubsystemResults')
    error('RCA_ShowSubsystemReview:MissingResults', ...
        'Could not resolve a valid RCA results struct. Run Vehicle_Detailed_Analysis first or pass RCA_Results explicitly.');
end
end

function sub = localFindSubsystem(results, subsystemKey)
sub = [];
requestedAliases = localSubsystemAliases(subsystemKey);
for iSub = 1:numel(results.SubsystemResults)
    candidateKey = upper(regexprep(string(results.SubsystemResults(iSub).Name), '[^A-Za-z0-9]', ''));
    candidateAliases = localSubsystemAliases(candidateKey);
    if any(candidateAliases == subsystemKey) || any(requestedAliases == candidateKey) || ...
            any(ismember(candidateAliases, requestedAliases)) || ...
            contains(candidateKey, subsystemKey) || contains(subsystemKey, candidateKey)
        sub = results.SubsystemResults(iSub);
        return;
    end
end
end

function aliases = localSubsystemAliases(subsystemKey)
aliases = string(subsystemKey);
normalized = upper(regexprep(string(subsystemKey), '[^A-Za-z0-9]', ''));
aliases(end + 1) = normalized;

switch normalized
    case "POWERTRAINCONTROLLER"
        aliases = [aliases, "POWERTRAINCONTROLLER", "POWERTRAIN", "PTCONTROLLER", "ANALYZEPOWERTRAINCONTROLLER"];
    case "ENVIRONMENT"
        aliases = [aliases, "ENVIRONMENT", "ANALYZEENVIRONMENT"];
    case "DRIVER"
        aliases = [aliases, "DRIVER", "ANALYZEDRIVER"];
    case "ELECTRICDRIVE"
        aliases = [aliases, "ELECTRICDRIVE", "EDRIVE", "MOTORDRIVE", "ANALYZEELECTRICDRIVE"];
    case "TRANSMISSION"
        aliases = [aliases, "TRANSMISSION", "GEARBOX", "ANALYZETRANSMISSION"];
    case "FINALDRIVE"
        aliases = [aliases, "FINALDRIVE", "DIFF", "AXLEDRIVE", "ANALYZEFINALDRIVE"];
    case "PNEUMATICBRAKESYSTEM"
        aliases = [aliases, "PNEUMATICBRAKESYSTEM", "PNEUMATICBRAKE", "BRAKESYSTEM", "ANALYZEPNEUMATICBRAKESYSTEM"];
end

aliases = unique(aliases);
end

function names = localAvailableSubsystemNames(results)
names = strings(0, 1);
if ~isfield(results, 'SubsystemResults') || isempty(results.SubsystemResults)
    names = "None";
    return;
end

for iSub = 1:numel(results.SubsystemResults)
    names(end + 1, 1) = string(results.SubsystemResults(iSub).Name); %#ok<AGROW>
end
names = unique(names);
end

function textValue = localSignalList(signalValue)
if isempty(signalValue)
    textValue = "None";
else
    textValue = strjoin(string(signalValue(:)'), ', ');
end
end

function localPrintTextBlock(lines)
lines = string(lines(:));
lines(lines == "") = [];
for iLine = 1:numel(lines)
    fprintf('  - %s\n', char(lines(iLine)));
end
end

function localOpenImageFigure(filePath, figureName)
try
    imageData = imread(char(filePath));
    fig = figure('Color', 'w', 'Name', figureName, 'NumberTitle', 'off');
    ax = axes('Parent', fig);
    image(ax, imageData);
    axis(ax, 'image');
    axis(ax, 'off');
    title(ax, figureName, 'Interpreter', 'none');
catch imageException
    warning('RCA_ShowSubsystemReview:FigureOpen', ...
        'Could not display figure %s: %s', char(filePath), imageException.message);
end
end

function files = localExistingFiles(fileList)
fileList = string(fileList(:));
mask = arrayfun(@(x) strlength(x) > 0 && isfile(char(x)), fileList);
files = fileList(mask);
end

function pretty = localPrettyName(nameValue)
pretty = regexprep(char(string(nameValue)), '([a-z])([A-Z])', '$1 $2');
pretty = strrep(pretty, '_', ' ');
pretty = strtrim(pretty);
if isempty(pretty)
    pretty = 'Subsystem';
end
end

function localPrintPlainTable(inputTable)
if ~istable(inputTable) || isempty(inputTable)
    return;
end

displayTable = localReorderDisplayColumns(inputTable);
headers = displayTable.Properties.VariableNames;
nRows = height(inputTable);
nCols = width(inputTable);

textData = cell(nRows, nCols);
colWidths = zeros(1, nCols);

for iCol = 1:nCols
    colWidths(iCol) = strlength(string(headers{iCol}));
end

for iRow = 1:nRows
    for iCol = 1:nCols
        textValue = localScalarToText(displayTable{iRow, iCol});
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

function displayTable = localReorderDisplayColumns(inputTable)
displayTable = inputTable;
requiredNames = ["SignalBasis", "StatusNote"];
if ~all(ismember(requiredNames, string(displayTable.Properties.VariableNames)))
    return;
end

currentNames = string(displayTable.Properties.VariableNames);
keepNames = currentNames(~ismember(currentNames, requiredNames));
displayOrder = [keepNames, "StatusNote", "SignalBasis"];
displayTable = displayTable(:, cellstr(displayOrder));
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
