% DIVe_sMP_Compare.m
% Compares nested sMP branches from two MAT files and exports the result to Excel.

clearvars;
clc;

selectedFiles = selectMatFilesFromGui(pwd);
if selectedFiles.Cancelled
    fprintf('MAT file selection cancelled. Script stopped.\n');
    return;
end

filePathA = selectedFiles.FilePathA;
filePathB = selectedFiles.FilePathB;

fprintf('Selected MAT File 1: %s\n', char(filePathA));
fprintf('Selected MAT File 2: %s\n', char(filePathB));
fprintf('\n');

statusBox = initStatusBox('Preparing sMP comparison...');
statusCleanup = onCleanup(@()closeStatusBox(statusBox));

fprintf('Loading sMP.phys, sMP.ctrl, sMP.human and sMP.bdry from the selected MAT files...\n');
updateStatusBox(statusBox, 0.10, 'Loading first MAT file...');
matDataA = loadSmpData(filePathA);
updateStatusBox(statusBox, 0.25, 'Loading second MAT file...');
matDataB = loadSmpData(filePathB);

updateStatusBox(statusBox, 0.35, 'Comparing sMP branches...');
comparisonCells = buildComparisonCells(matDataA, matDataB, statusBox);
comparisonTable = cell2table(comparisonCells(2:end, :), ...
    'VariableNames', buildComparisonTableVariableNames(comparisonCells(1, :)));

updateStatusBox(statusBox, 0.80, 'Creating results folder...');
resultsDir = createSmpResultsFolder(filePathA);
outputWorkbookPath = buildWorkbookPath(resultsDir, filePathA, filePathB);
updateStatusBox(statusBox, 0.88, 'Writing Excel report...');
writecell(comparisonCells, outputWorkbookPath, 'Sheet', 'sMP_Compare', 'Range', 'A1');
updateStatusBox(statusBox, 0.94, 'Applying Excel formatting...');
formatComparisonWorkbook(outputWorkbookPath, comparisonCells);
updateStatusBox(statusBox, 1.00, 'Comparison complete.');
pause(0.3);
closeStatusBox(statusBox);
statusCleanup = [];

assignin('base', 'diveSmpComparisonTable', comparisonTable);
assignin('base', 'diveSmpComparisonCells', comparisonCells);
assignin('base', 'diveSmpComparisonWorkbookPath', outputWorkbookPath);
assignin('base', 'diveSmpComparisonResultsDir', resultsDir);

fprintf('\nExcel report location: %s\n', outputWorkbookPath);
printWorkbookLink(outputWorkbookPath);

function h = initStatusBox(initialMessage)
h = [];
if ~usejava('desktop')
    return;
end

stale = findall(0, 'Type', 'figure', 'Tag', 'ExecutionStatusWaitbar');
for iFig = 1:numel(stale)
    try
        delete(stale(iFig));
    catch
    end
end

try
    h = waitbar(0, normalizeStatusMessage(initialMessage), ...
        'Name', 'DIVe sMP Compare', ...
        'CreateCancelBtn', '');
    set(h, 'Tag', 'ExecutionStatusWaitbar');
    setappdata(h, 'StartTic', tic);
    configureWaitbarTextInterpreter(h);
catch
    h = [];
end
end

function updateStatusBox(h, fraction, messageText)
if isempty(h) || ~ishandle(h)
    return;
end

f = max(0, min(1, double(fraction)));
try
    waitbar(f, h, buildStatusDisplayMessage(h, messageText, f));
    configureWaitbarTextInterpreter(h);
    drawnow;
catch
end
end

function closeStatusBox(h)
if isempty(h) || ~ishandle(h)
    return;
end
try
    delete(h);
catch
end
end

function out = normalizeStatusMessage(messageText)
out = char(string(messageText));
if strlength(string(out)) == 0
    out = 'Working...';
end
end

function out = buildStatusDisplayMessage(h, messageText, fraction)
baseMessage = normalizeStatusMessage(messageText);
startTic = [];
try
    startTic = getappdata(h, 'StartTic');
catch
end

if isempty(startTic) || ~isscalar(fraction) || fraction <= 0
    out = baseMessage;
    return;
end

elapsedSeconds = toc(startTic);
if fraction >= 1
    remainingSeconds = 0;
else
    remainingSeconds = max(0, elapsedSeconds * (1 - fraction) / max(fraction, eps));
end

out = sprintf('%s\nApprox. time remaining: %s', ...
    baseMessage, formatDurationText(remainingSeconds));
end

function out = formatDurationText(totalSeconds)
totalSeconds = max(0, round(double(totalSeconds)));
hoursPart = floor(totalSeconds / 3600);
minutesPart = floor(mod(totalSeconds, 3600) / 60);
secondsPart = mod(totalSeconds, 60);

if hoursPart > 0
    out = sprintf('%02d:%02d:%02d', hoursPart, minutesPart, secondsPart);
else
    out = sprintf('%02d:%02d', minutesPart, secondsPart);
end
end

function configureWaitbarTextInterpreter(h)
if isempty(h) || ~ishandle(h)
    return;
end

try
    txt = findall(h, 'Type', 'Text');
    set(txt, 'Interpreter', 'none');
catch
end
end

function selection = selectMatFilesFromGui(startPath)
selection = struct('FilePathA', "", 'FilePathB', "", 'Cancelled', true);

if ~usejava('desktop')
    warning('DIVe_sMP_Compare:NoDesktop', ...
        'MATLAB desktop is not available. Falling back to file dialogs.');
    selection.FilePathA = selectSingleMatFile('Select MAT File 1', startPath);
    if strlength(selection.FilePathA) == 0
        return;
    end
    selection.FilePathB = selectSingleMatFile('Select MAT File 2', fileparts(char(selection.FilePathA)));
    if strlength(selection.FilePathB) == 0
        selection.FilePathA = "";
        return;
    end
    selection.Cancelled = false;
    return;
end

dialogState = struct( ...
    'FilePathA', "", ...
    'FilePathB', "", ...
    'Cancelled', true);

fig = uifigure( ...
    'Name', 'DIVe sMP Compare', ...
    'Position', [100 100 900 220], ...
    'Resize', 'off');

fig.CloseRequestFcn = @onCancel;

uilabel(fig, ...
    'Text', 'Mat File 1', ...
    'Position', [40 145 100 22], ...
    'HorizontalAlignment', 'left');

fieldA = uieditfield(fig, 'text', ...
    'Position', [140 140 620 32], ...
    'Editable', 'off', ...
    'Value', '');

uibutton(fig, ...
    'Text', 'Browse...', ...
    'Position', [775 140 90 32], ...
    'ButtonPushedFcn', @onBrowseA);

uilabel(fig, ...
    'Text', 'Mat File 2', ...
    'Position', [40 85 100 22], ...
    'HorizontalAlignment', 'left');

fieldB = uieditfield(fig, 'text', ...
    'Position', [140 80 620 32], ...
    'Editable', 'off', ...
    'Value', '');

uibutton(fig, ...
    'Text', 'Browse...', ...
    'Position', [775 80 90 32], ...
    'ButtonPushedFcn', @onBrowseB);

continueBtn = uibutton(fig, ...
    'Text', 'Continue', ...
    'Position', [730 25 135 36], ...
    'Enable', 'off', ...
    'ButtonPushedFcn', @onContinue);

uiwait(fig);

selection = dialogState;
if isvalid(fig)
    delete(fig);
end

    function onBrowseA(~, ~)
        selectedPath = selectSingleMatFile('Select MAT File 1', resolveStartFolder(dialogState.FilePathA, startPath));
        if strlength(selectedPath) == 0
            return;
        end
        dialogState.FilePathA = selectedPath;
        fieldA.Value = char(selectedPath);
        updateContinueState();
    end

    function onBrowseB(~, ~)
        fallbackPath = dialogState.FilePathA;
        if strlength(fallbackPath) == 0
            fallbackPath = startPath;
        else
            fallbackPath = fileparts(char(fallbackPath));
        end
        selectedPath = selectSingleMatFile('Select MAT File 2', resolveStartFolder(dialogState.FilePathB, fallbackPath));
        if strlength(selectedPath) == 0
            return;
        end
        dialogState.FilePathB = selectedPath;
        fieldB.Value = char(selectedPath);
        updateContinueState();
    end

    function updateContinueState()
        hasBothFiles = strlength(dialogState.FilePathA) > 0 && strlength(dialogState.FilePathB) > 0;
        if hasBothFiles
            continueBtn.Enable = 'on';
        else
            continueBtn.Enable = 'off';
        end
    end

    function onContinue(~, ~)
        dialogState.Cancelled = false;
        uiresume(fig);
    end

    function onCancel(src, ~)
        dialogState.Cancelled = true;
        uiresume(src);
    end
end

function startFolder = resolveStartFolder(currentSelection, fallbackPath)
if strlength(string(currentSelection)) > 0 && isfile(char(string(currentSelection)))
    startFolder = fileparts(char(string(currentSelection)));
elseif strlength(string(fallbackPath)) > 0
    startFolder = char(string(fallbackPath));
else
    startFolder = pwd;
end
end

function filePath = selectSingleMatFile(dialogTitle, startPath)
[fileName, folderPath] = uigetfile('*.mat', dialogTitle, startPath);
if isequal(fileName, 0)
    filePath = "";
    return;
end
filePath = string(fullfile(folderPath, fileName));
end

function matData = loadSmpData(filePath)
requiredSections = getRequiredSmpSections();
loadedData = load(char(filePath), 'sMP');
hasSmp = isfield(loadedData, 'sMP') && isstruct(loadedData.sMP);

[~, fileLabel, ~] = fileparts(char(filePath));
matData = struct( ...
    'FilePath', string(filePath), ...
    'FileLabel', string(fileLabel), ...
    'HasSmp', hasSmp, ...
    'Value', struct(), ...
    'AvailableSections', strings(0, 1));

if hasSmp
    matData.Value = loadedData.sMP;
    presentMask = isfield(loadedData.sMP, cellstr(requiredSections));
    matData.AvailableSections = requiredSections(presentMask);
else
    warning('DIVe_sMP_Compare:MissingSmp', ...
        'File "%s" does not contain sMP. Missing values will be reported as NA.', ...
        char(filePath));
end
end

function comparisonCells = buildComparisonCells(matDataA, matDataB, statusBox)
requiredSections = getRequiredSmpSections();
mapA = containers.Map('KeyType', 'char', 'ValueType', 'any');
mapB = containers.Map('KeyType', 'char', 'ValueType', 'any');
pathsA = strings(0, 1);
pathsB = strings(0, 1);
missingSectionPaths = strings(0, 1);

if matDataA.HasSmp
    [mapA, pathsA] = buildSmpSectionMap(matDataA.Value, requiredSections);
end

if matDataB.HasSmp
    [mapB, pathsB] = buildSmpSectionMap(matDataB.Value, requiredSections);
end

allPaths = union(pathsA, pathsB, 'stable');
for iSection = 1:numel(requiredSections)
    sectionPath = "sMP." + requiredSections(iSection);
    if ~any(pathsA == sectionPath) && ~any(pathsB == sectionPath) && ...
            ~any(startsWith(pathsA, sectionPath + ".")) && ~any(startsWith(pathsB, sectionPath + "."))
        missingSectionPaths(end + 1, 1) = sectionPath; %#ok<AGROW>
    end
end

allPaths = union(missingSectionPaths, allPaths, 'stable');
if isempty(allPaths)
    allPaths = "sMP";
end

pathPartsPerRow = cell(numel(allPaths), 1);
maxPathDepth = 1;
for iPath = 1:numel(allPaths)
    pathPartsPerRow{iPath} = splitPathIntoColumns(allPaths(iPath));
    maxPathDepth = max(maxPathDepth, numel(pathPartsPerRow{iPath}));
end

comparisonCells = cell(numel(allPaths) + 1, maxPathDepth + 3);
comparisonCells(1, 1:maxPathDepth) = buildPathHeaderRow(maxPathDepth);
comparisonCells(1, maxPathDepth + 1:end) = { ...
    char(matDataA.FileLabel), ...
    char(matDataB.FileLabel), ...
    'Status'};

progressStep = max(1, ceil(numel(allPaths) / 100));
for iPath = 1:numel(allPaths)
    currentPath = char(allPaths(iPath));
    [hasValueA, valueA] = getMapValue(mapA, currentPath);
    [hasValueB, valueB] = getMapValue(mapB, currentPath);

    currentPathParts = pathPartsPerRow{iPath};
    for iPart = 1:numel(currentPathParts)
        comparisonCells{iPath + 1, iPart} = currentPathParts{iPart};
    end
    comparisonCells{iPath + 1, maxPathDepth + 1} = formatValueForExcel(hasValueA, valueA);
    comparisonCells{iPath + 1, maxPathDepth + 2} = formatValueForExcel(hasValueB, valueB);
    comparisonCells{iPath + 1, maxPathDepth + 3} = hasValueA && hasValueB && isequaln(valueA, valueB);

    if iPath == 1 || iPath == numel(allPaths) || mod(iPath, progressStep) == 0
        progressValue = 0.35 + 0.40 * iPath / max(1, numel(allPaths));
        updateStatusBox(statusBox, progressValue, ...
            sprintf('Comparing variable %d of %d...', iPath, numel(allPaths)));
    end
end
end

function [valueMap, orderedPaths] = buildSmpSectionMap(smpValue, requiredSections)
valueMap = containers.Map('KeyType', 'char', 'ValueType', 'any');
orderedPaths = strings(0, 1);

for iSection = 1:numel(requiredSections)
    sectionName = char(requiredSections(iSection));
    if ~isfield(smpValue, sectionName)
        continue;
    end

    entries = flattenValueTree(['sMP.' sectionName], smpValue.(sectionName));
    [sectionMap, sectionPaths] = buildValueMap(entries);
    for iPath = 1:numel(sectionPaths)
        currentPath = char(sectionPaths(iPath));
        valueMap(currentPath) = sectionMap(currentPath);
    end
    orderedPaths = [orderedPaths; sectionPaths]; %#ok<AGROW>
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

function pathParts = splitPathIntoColumns(pathText)
pathText = char(string(pathText));
if strlength(string(pathText)) == 0
    pathParts = {''};
    return;
end
pathParts = strsplit(pathText, '.');
end

function headerRow = buildPathHeaderRow(maxPathDepth)
headerRow = cell(1, maxPathDepth);
fixedHeaders = {'sMP Variable', 'Context', 'Species'};
fixedCount = min(maxPathDepth, numel(fixedHeaders));
headerRow(1:fixedCount) = fixedHeaders(1:fixedCount);

for iCol = 4:maxPathDepth
    headerRow{iCol} = sprintf('Var_%d', iCol - 3);
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
        outText = double(value);
    elseif islogical(value)
        outText = logical(value);
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

function requiredSections = getRequiredSmpSections()
requiredSections = ["phys"; "ctrl"; "human"; "bdry"];
end

function resultsDir = createSmpResultsFolder(filePathA)
baseDir = fileparts(char(filePathA));
timeStamp = char(datetime('now', 'Format', 'yyyyMMdd_HHmmss'));
folderName = sprintf('%s_sMP_Compare_Results', timeStamp);
resultsDir = fullfile(baseDir, folderName);

if ~isfolder(resultsDir)
    mkdir(resultsDir);
    return;
end

suffix = 1;
while isfolder(resultsDir)
    resultsDir = fullfile(baseDir, sprintf('%s_%02d', folderName, suffix));
    suffix = suffix + 1;
end
mkdir(resultsDir);
end

function workbookPath = buildWorkbookPath(resultsDir, filePathA, filePathB)
[~, nameA, ~] = fileparts(char(filePathA));
[~, nameB, ~] = fileparts(char(filePathB));
defaultFileName = sprintf('sMP_compare_%s_vs_%s.xlsx', nameA, nameB);
workbookPath = fullfile(resultsDir, defaultFileName);
end

function formatComparisonWorkbook(workbookPath, comparisonCells)
if strlength(string(workbookPath)) == 0 || ~isfile(char(string(workbookPath)))
    return;
end

rowCount = size(comparisonCells, 1);
colCount = size(comparisonCells, 2);
if rowCount < 1 || colCount < 1
    return;
end

excelApp = [];
wb = [];
try
    excelApp = actxserver('Excel.Application');
    excelApp.DisplayAlerts = false;
    excelApp.Visible = false;
    wb = excelApp.Workbooks.Open(char(string(workbookPath)), false, false);
    ws = wb.Worksheets.Item('sMP_Compare');

    headerRange = getExcelRange(ws, 1, 1, 1, colCount);
    headerRange.Font.Bold = true;
    headerRange.Font.Color = excelRgb(255, 255, 255);
    headerRange.Interior.Color = excelRgb(31, 78, 121);
    invoke(headerRange, 'AutoFilter');

    usedRange = getExcelRange(ws, 1, 1, rowCount, colCount);
    usedRange.Columns.AutoFit;
    setComparisonColumnWidths(ws, comparisonCells);
    ws.Activate;
    excelApp.ActiveWindow.SplitColumn = 0;
    excelApp.ActiveWindow.SplitRow = 1;
    excelApp.ActiveWindow.FreezePanes = true;

    if rowCount >= 2
        statusCol = colCount;
        for iRow = 2:rowCount
            statusCell = getExcelRange(ws, iRow, statusCol, iRow, statusCol);
            statusValue = comparisonCells{iRow, statusCol};
            if islogical(statusValue) && isscalar(statusValue) && statusValue
                statusCell.Font.Bold = false;
                statusCell.Font.Color = excelRgb(255, 255, 255);
                statusCell.Interior.Color = excelRgb(0, 97, 0);
            elseif islogical(statusValue) && isscalar(statusValue) && ~statusValue
                statusCell.Font.Bold = true;
                statusCell.Font.Color = excelRgb(255, 255, 255);
                statusCell.Interior.Color = excelRgb(192, 80, 77);
            end
        end
    end

    wb.Save;
catch ME
    warning('DIVe_sMP_Compare:ExcelFormattingFailed', ...
        'Workbook was created but Excel formatting could not be applied: %s', ...
        ME.message);
end

cleanupExcelSession(excelApp, wb);
end

function variableNames = buildComparisonTableVariableNames(headerRow)
headerText = strings(1, numel(headerRow));
for iCol = 1:numel(headerRow)
    headerValue = headerRow{iCol};
    if isstring(headerValue) || ischar(headerValue)
        headerText(iCol) = string(headerValue);
    else
        headerText(iCol) = "Column_" + iCol;
    end
end
variableNames = matlab.lang.makeUniqueStrings(matlab.lang.makeValidName(cellstr(headerText)));
end

function printWorkbookLink(workbookPath)
if strlength(string(workbookPath)) == 0 || ~isfile(char(string(workbookPath)))
    return;
end

[~, workbookName, workbookExt] = fileparts(char(string(workbookPath)));
workbookCmd = sprintf('matlab:winopen(''%s'')', escapeForMatlabCharLiteral(workbookPath));
fprintf('\n');
fprintf('Excel Report Link: <a href="%s">%s%s</a>\n', workbookCmd, workbookName, workbookExt);
fprintf('\n');
end

function out = escapeForMatlabCharLiteral(inText)
out = strrep(char(string(inText)), '''', '''''');
end

function colorValue = excelRgb(redValue, greenValue, blueValue)
colorValue = double(redValue) + 256 * double(greenValue) + 65536 * double(blueValue);
end

function rangeObj = getExcelRange(sheetObj, startRow, startCol, endRow, endCol)
if nargin < 5
    endRow = startRow;
    endCol = startCol;
end
rangeAddress = sprintf('%s:%s', xlA1(startRow, startCol), xlA1(endRow, endCol));
rangeObj = get(sheetObj, 'Range', rangeAddress);
end

function setComparisonColumnWidths(sheetObj, comparisonCells)
colCount = size(comparisonCells, 2);
if colCount < 3
    return;
end

for iCol = 1:colCount
    headerValue = "";
    if ~isempty(comparisonCells{1, iCol}) && (ischar(comparisonCells{1, iCol}) || isstring(comparisonCells{1, iCol}))
        headerValue = string(comparisonCells{1, iCol});
    end

    if startsWith(headerValue, "Var_") || iCol == colCount - 2 || iCol == colCount - 1
        columnRange = getExcelRange(sheetObj, 1, iCol, size(comparisonCells, 1), iCol);
        columnRange.ColumnWidth = 40;
    end
end
end

function addr = xlA1(rowIdx, colIdx)
addr = [xlColName(colIdx) num2str(rowIdx)];
end

function colName = xlColName(colIdx)
colIdx = double(colIdx);
if colIdx < 1
    colName = 'A';
    return;
end

colName = '';
while colIdx > 0
    remVal = mod(colIdx - 1, 26);
    colName = [char(65 + remVal) colName]; %#ok<AGROW>
    colIdx = floor((colIdx - 1) / 26);
end
end

function cleanupExcelSession(excelApp, wb)
try
    if ~isempty(wb)
        wb.Close(false);
    end
catch
end
try
    if ~isempty(excelApp)
        excelApp.Quit;
    end
catch
end
try
    if ~isempty(excelApp)
        delete(excelApp);
    end
catch
end
end
