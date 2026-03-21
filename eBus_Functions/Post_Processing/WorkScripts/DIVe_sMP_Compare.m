% DIVe_sMP_Compare.m
% Compares nested sMP.phys values from two MAT files and exports the result to Excel.

clearvars;
clc;

filePathA = selectSingleMatFile('Select the first MAT file', pwd);
if strlength(filePathA) == 0
    fprintf('First MAT file selection cancelled. Script stopped.\n');
    return;
end

filePathB = selectSingleMatFile('Select the second MAT file', fileparts(char(filePathA)));
if strlength(filePathB) == 0
    fprintf('Second MAT file selection cancelled. Script stopped.\n');
    return;
end

fprintf('Loading sMP.phys from the selected MAT files...\n');
matDataA = loadSmpPhys(filePathA);
matDataB = loadSmpPhys(filePathB);

comparisonCells = buildComparisonCells(matDataA, matDataB);
comparisonTable = cell2table(comparisonCells(2:end, :), ...
    'VariableNames', {'SMPVariable', 'FirstFileValue', 'SecondFileValue', 'Status'});

defaultWorkbookPath = buildDefaultWorkbookPath(filePathA, filePathB);
[outputFileName, outputFolder] = uiputfile('*.xlsx', ...
    'Save sMP comparison report as', defaultWorkbookPath);

if isequal(outputFileName, 0)
    fprintf('Excel export cancelled. Comparison table is available in the MATLAB workspace.\n');
    assignin('base', 'diveSmpComparisonTable', comparisonTable);
    assignin('base', 'diveSmpComparisonCells', comparisonCells);
    return;
end

outputWorkbookPath = fullfile(outputFolder, outputFileName);
writecell(comparisonCells, outputWorkbookPath, 'Sheet', 'sMP_Compare', 'Range', 'A1');

assignin('base', 'diveSmpComparisonTable', comparisonTable);
assignin('base', 'diveSmpComparisonCells', comparisonCells);
assignin('base', 'diveSmpComparisonWorkbookPath', outputWorkbookPath);

fprintf('sMP comparison workbook created: %s\n', outputWorkbookPath);
try
    winopen(outputWorkbookPath);
catch
end

function filePath = selectSingleMatFile(dialogTitle, startPath)
[fileName, folderPath] = uigetfile('*.mat', dialogTitle, startPath);
if isequal(fileName, 0)
    filePath = "";
    return;
end
filePath = string(fullfile(folderPath, fileName));
end

function matData = loadSmpPhys(filePath)
loadedData = load(char(filePath), 'sMP');
hasSmpPhys = isfield(loadedData, 'sMP') && isstruct(loadedData.sMP) && isfield(loadedData.sMP, 'phys');

[~, fileLabel, ~] = fileparts(char(filePath));
matData = struct( ...
    'FilePath', string(filePath), ...
    'FileLabel', string(fileLabel), ...
    'HasSmpPhys', hasSmpPhys, ...
    'Value', []);

if hasSmpPhys
    matData.Value = loadedData.sMP.phys;
else
    warning('DIVe_sMP_Compare:MissingSmpPhys', ...
        'File "%s" does not contain sMP.phys. Missing values will be reported as NA.', ...
        char(filePath));
end
end

function comparisonCells = buildComparisonCells(matDataA, matDataB)
mapA = containers.Map('KeyType', 'char', 'ValueType', 'any');
mapB = containers.Map('KeyType', 'char', 'ValueType', 'any');
pathsA = strings(0, 1);
pathsB = strings(0, 1);

if matDataA.HasSmpPhys
    entriesA = flattenValueTree('sMP.phys', matDataA.Value);
    [mapA, pathsA] = buildValueMap(entriesA);
end

if matDataB.HasSmpPhys
    entriesB = flattenValueTree('sMP.phys', matDataB.Value);
    [mapB, pathsB] = buildValueMap(entriesB);
end

allPaths = union(pathsA, pathsB, 'stable');
if isempty(allPaths)
    allPaths = "sMP.phys";
end

comparisonCells = cell(numel(allPaths) + 1, 4);
comparisonCells(1, :) = { ...
    'sMP Variable', ...
    char(matDataA.FileLabel), ...
    char(matDataB.FileLabel), ...
    'Status'};

for iPath = 1:numel(allPaths)
    currentPath = char(allPaths(iPath));
    [hasValueA, valueA] = getMapValue(mapA, currentPath);
    [hasValueB, valueB] = getMapValue(mapB, currentPath);

    comparisonCells{iPath + 1, 1} = currentPath;
    comparisonCells{iPath + 1, 2} = formatValueForExcel(hasValueA, valueA);
    comparisonCells{iPath + 1, 3} = formatValueForExcel(hasValueB, valueB);
    comparisonCells{iPath + 1, 4} = hasValueA && hasValueB && isequaln(valueA, valueB);
end
end

function entries = flattenValueTree(currentPath, value)
entries = struct('Path', {}, 'Value', {});

if isstruct(value)
    if isempty(value)
        entries(1).Path = currentPath;
        entries(1).Value = value;
        return;
    end

    if isscalar(value)
        fieldNames = fieldnames(value);
        if isempty(fieldNames)
            entries(1).Path = currentPath;
            entries(1).Value = value;
            return;
        end

        for iField = 1:numel(fieldNames)
            fieldName = fieldNames{iField};
            childPath = sprintf('%s.%s', currentPath, fieldName);
            childEntries = flattenValueTree(childPath, value.(fieldName));
            entries = [entries, childEntries]; %#ok<AGROW>
        end
        return;
    end

    for iElement = 1:numel(value)
        indexedPath = sprintf('%s(%d)', currentPath, iElement);
        childEntries = flattenValueTree(indexedPath, value(iElement));
        entries = [entries, childEntries]; %#ok<AGROW>
    end
    return;
end

entries(1).Path = currentPath;
entries(1).Value = value;
end

function [valueMap, orderedPaths] = buildValueMap(entries)
valueMap = containers.Map('KeyType', 'char', 'ValueType', 'any');
orderedPaths = strings(numel(entries), 1);
for iEntry = 1:numel(entries)
    valueMap(entries(iEntry).Path) = entries(iEntry).Value;
    orderedPaths(iEntry) = string(entries(iEntry).Path);
end
end

function [hasValue, value] = getMapValue(valueMap, key)
hasValue = isKey(valueMap, key);
if hasValue
    value = valueMap(key);
else
    value = [];
end
end

function outText = formatValueForExcel(hasValue, value)
if ~hasValue
    outText = 'NA';
    return;
end

if isstring(value)
    outText = formatStringValue(value);
    return;
end

if ischar(value)
    if isrow(value)
        outText = value;
    else
        outText = sizeToText(size(value));
    end
    return;
end

if isnumeric(value) || islogical(value)
    outText = formatNumericLikeValue(value);
    return;
end

if iscell(value)
    outText = formatCellValue(value);
    return;
end

if istable(value)
    outText = sprintf('table %s', sizeToText(size(value)));
    return;
end

if isa(value, 'datetime') || isa(value, 'duration') || isa(value, 'calendarDuration')
    outText = formatDisplayArray(string(value));
    return;
end

if isstruct(value)
    outText = sprintf('struct %s', sizeToText(size(value)));
    return;
end

valueSize = size(value);
if isscalar(value)
    try
        outText = char(string(value));
    catch
        outText = sprintf('%s scalar', class(value));
    end
elseif isvector(value) && numel(value) <= 5
    try
        outText = formatDisplayArray(string(value));
    catch
        outText = sprintf('%s %s', class(value), sizeToText(valueSize));
    end
else
    outText = sprintf('%s %s', class(value), sizeToText(valueSize));
end
end

function outText = formatStringValue(value)
if isscalar(value)
    outText = char(value);
    return;
end

if isvector(value) && numel(value) <= 5
    outText = formatDisplayArray(value);
else
    outText = sizeToText(size(value));
end
end

function outText = formatNumericLikeValue(value)
valueSize = size(value);

if isempty(value)
    outText = sizeToText(valueSize);
    return;
end

if isscalar(value)
    if isnumeric(value)
        outText = num2str(value, 15);
    else
        outText = mat2str(value);
    end
    return;
end

if isvector(value)
    if numel(value) <= 5
        outText = mat2str(value);
    else
        outText = sizeToText(valueSize);
    end
    return;
end

outText = sizeToText(valueSize);
end

function outText = formatCellValue(value)
valueSize = size(value);
if isempty(value)
    outText = sizeToText(valueSize);
    return;
end

if isvector(value) && numel(value) <= 5
    renderedValues = strings(1, numel(value));
    for iValue = 1:numel(value)
        renderedValues(iValue) = string(formatCellElement(value{iValue}));
    end
    outText = sprintf('{%s}', strjoin(cellstr(renderedValues), ', '));
else
    outText = sizeToText(valueSize);
end
end

function outText = formatCellElement(value)
if isstring(value)
    outText = formatStringValue(value);
elseif ischar(value) && isrow(value)
    outText = value;
elseif isnumeric(value) || islogical(value)
    outText = formatNumericLikeValue(value);
elseif iscell(value)
    outText = sizeToText(size(value));
elseif isstruct(value)
    outText = sprintf('struct %s', sizeToText(size(value)));
else
    try
        outText = char(string(value));
    catch
        outText = sprintf('%s %s', class(value), sizeToText(size(value)));
    end
end
end

function outText = formatDisplayArray(value)
value = string(value);
if isempty(value)
    outText = sizeToText(size(value));
    return;
end
outText = sprintf('[%s]', strjoin(cellstr(value(:).'), ', '));
end

function outText = sizeToText(dimensions)
dimensionStrings = arrayfun(@num2str, dimensions, 'UniformOutput', false);
outText = strjoin(dimensionStrings, 'x');
end

function workbookPath = buildDefaultWorkbookPath(filePathA, filePathB)
[~, nameA, ~] = fileparts(char(filePathA));
[~, nameB, ~] = fileparts(char(filePathB));
defaultFileName = sprintf('sMP_phys_compare_%s_vs_%s.xlsx', nameA, nameB);
workbookPath = fullfile(fileparts(char(filePathA)), defaultFileName);
end
