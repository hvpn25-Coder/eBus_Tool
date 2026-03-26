function CreateDIVeDatasetVariant
% CREATEDIVEDATASETVARIANT
% 1) Ask user to select a folder.
% 2) Read Veh_Param.m variable assignments.
% 3) Show GUI table with Variable, Value, and Edit_1..Edit_10.
% 4) On Save: for each non-empty Edit_N column, duplicate folder with
%    timestamp + EN suffix, update Veh_Param.m, and update XML name/content.

selectedFolder = uigetdir(pwd, 'Select folder containing parameter .m file');
if isequal(selectedFolder, 0)
    return;
end

[selectedMFileName, selectedMFilePath] = pickTargetMFile(selectedFolder);
if isempty(selectedMFilePath)
    return;
end

[entries, fileText] = parseVehParamFile(selectedMFilePath);
if isempty(entries)
    errordlg(sprintf('No variable assignments were found in %s.', selectedMFileName), 'Parse Error');
    return;
end

varNames = {entries.name}.';
varValues = {entries.value}.';
numEditColumns = 10;
editColumnHeaders = arrayfun(@(n)sprintf('Edit_%d', n), 1:numEditColumns, 'UniformOutput', false);
tableData = [varNames, varValues, repmat({''}, numel(entries), numEditColumns)];

fig = uifigure('Name', 'Veh Param Editor', 'Position', [100 100 900 620]);
fig.AutoResizeChildren = 'off';

tbl = uitable( ...
    fig, ...
    'Data', tableData, ...
    'ColumnName', [{'Variable', 'Value'}, editColumnHeaders], ...
    'RowName', {}, ...
    'ColumnEditable', [false false true(1, numEditColumns)], ...
    'Position', [20 100 720 340], ...
    'CellEditCallback', @(~, event)handleTableEdit(fig, event));

filterLabel = uilabel( ...
    fig, ...
    'Text', 'Filter Variable:', ...
    'Position', [20 470 100 22]);

filterField = uieditfield( ...
    fig, ...
    'text', ...
    'Placeholder', 'Type variable name to filter', ...
    'Position', [125 470 360 22], ...
    'ValueChangedFcn', @(src, ~)applyVariableFilter(fig, tbl, src.Value), ...
    'ValueChangingFcn', @(~, event)applyVariableFilter(fig, tbl, event.Value));

resetBtn = uibutton( ...
    fig, ...
    'Text', 'Reset Filter', ...
    'Position', [495 470 105 22], ...
    'ButtonPushedFcn', @(~, ~)resetVariableFilter(fig, tbl, filterField));

headerLabel = uilabel( ...
    fig, ...
    'Text', 'Editable Column Names:', ...
    'Position', [20 445 170 18]);

headerFields = gobjects(1, numEditColumns);
for idx = 1:numEditColumns
    headerFields(idx) = uieditfield( ...
        fig, ...
        'text', ...
        'Value', editColumnHeaders{idx}, ...
        'Position', [20 420 60 22], ...
        'ValueChangedFcn', @(src, ~)handleHeaderRename(fig, tbl, src, idx));
end

newVarLabel = uilabel( ...
    fig, ...
    'Text', 'New Variable:', ...
    'Position', [20 60 85 22]);

newVarField = uieditfield( ...
    fig, ...
    'text', ...
    'Placeholder', 'e.g., New_Param or Veh.New.Field', ...
    'Position', [105 60 230 22]);

newValLabel = uilabel( ...
    fig, ...
    'Text', 'Default Value:', ...
    'Position', [345 60 85 22]);

newValField = uieditfield( ...
    fig, ...
    'text', ...
    'Placeholder', 'e.g., 123 or struct(...)', ...
    'Position', [430 60 150 22]);

addRowBtn = uibutton( ...
    fig, ...
    'Text', 'Add Row', ...
    'Position', [590 60 90 24], ...
    'ButtonPushedFcn', @(~, ~)addVariableRow(fig, tbl, filterField, newVarField, newValField, numEditColumns));

nameModeGroup = uibuttongroup( ...
    fig, ...
    'Title', 'Folder Naming', ...
    'Position', [20 5 360 50], ...
    'SelectionChangedFcn', @(src, event)handleNamingModeChange(fig, src, event));

uiradiobutton( ...
    nameModeGroup, ...
    'Text', 'Date-Time Stamp', ...
    'Tag', 'timestamp', ...
    'Position', [10 5 140 18]);

rbHeader = uiradiobutton( ...
    nameModeGroup, ...
    'Text', 'Header Name (Editable)', ...
    'Tag', 'header', ...
    'Position', [160 5 170 18]);
rbHeader.Value = false;

saveBtn = uibutton( ...
    fig, ...
    'Text', 'Create DIVe Dataset', ...
    'Position', [570 20 160 30], ...
    'ButtonPushedFcn', @(~, ~)saveChanges(fileText, selectedFolder, selectedMFileName, fig, numEditColumns));

outputPanel = uipanel( ...
    fig, ...
    'Title', 'Output', ...
    'Position', [20 5 720 90]);

outputLabel = uilabel( ...
    outputPanel, ...
    'Text', 'Open the parent folder that contains all created datasets.', ...
    'Position', [10 34 700 22]);

outputLink = uihyperlink( ...
    outputPanel, ...
    'Text', 'Open Folder', ...
    'Position', [10 6 200 18], ...
    'Visible', 'off', ...
    'HyperlinkClickedFcn', @(~, ~)openSelectedOutputFolder(fig));

fig.UserData = struct( ...
    'allTableData', {tableData}, ...
    'visibleRowIdx', (1:size(tableData, 1)).', ...
    'entries', entries, ...
    'editHeaderNames', string(editColumnHeaders), ...
    'nameMode', "timestamp", ...
    'outputFolderPath', "", ...
    'numEditColumns', numEditColumns, ...
    'ui', struct( ...
        'tbl', tbl, ...
        'filterLabel', filterLabel, ...
        'filterField', filterField, ...
        'resetBtn', resetBtn, ...
        'headerLabel', headerLabel, ...
        'headerFields', {headerFields}, ...
        'newVarLabel', newVarLabel, ...
        'newVarField', newVarField, ...
        'newValLabel', newValLabel, ...
        'newValField', newValField, ...
        'addRowBtn', addRowBtn, ...
        'nameModeGroup', nameModeGroup, ...
        'saveBtn', saveBtn, ...
        'outputPanel', outputPanel, ...
        'outputLabel', outputLabel, ...
        'outputLink', outputLink));

layoutVehParamUI(fig);
fig.SizeChangedFcn = @(src, ~)layoutVehParamUI(src);
end

function [mFileName, mFilePath] = pickTargetMFile(selectedFolder)
mFileName = '';
mFilePath = '';

mFiles = dir(fullfile(selectedFolder, '*.m'));
if isempty(mFiles)
    errordlg('No .m file was found in the selected folder.', 'Missing File');
    return;
end

thisScriptName = [mfilename('name'), '.m'];
candidateNames = {mFiles.name};
candidateNames = candidateNames(~strcmpi(candidateNames, thisScriptName));
if isempty(candidateNames)
    candidateNames = {mFiles.name};
end

if isscalar(candidateNames)
    mFileName = candidateNames{1};
    mFilePath = fullfile(selectedFolder, mFileName);
    return;
end

[idx, ok] = listdlg( ...
    'ListString', candidateNames, ...
    'SelectionMode', 'single', ...
    'PromptString', 'Select the parameter .m file to process:', ...
    'Name', 'Choose .m file');
if ~ok || isempty(idx)
    return;
end

mFileName = candidateNames{idx};
mFilePath = fullfile(selectedFolder, mFileName);
end

function applyVariableFilter(fig, tbl, filterText)
state = fig.UserData;
allData = state.allTableData;

if isempty(strtrim(string(filterText)))
    visibleIdx = (1:size(allData, 1)).';
else
    varNames = string(allData(:, 1));
    visibleIdx = find(contains(lower(varNames), lower(string(filterText))));
end

state.visibleRowIdx = visibleIdx;
fig.UserData = state;
tbl.Data = allData(visibleIdx, :);
end

function resetVariableFilter(fig, tbl, filterField)
filterField.Value = '';
applyVariableFilter(fig, tbl, '');
end

function handleHeaderRename(fig, tbl, src, headerIndex)
newHeader = strtrim(string(src.Value));
if strlength(newHeader) == 0
    newHeader = "Edit_" + string(headerIndex);
end

state = fig.UserData;
state.editHeaderNames(headerIndex) = newHeader;
fig.UserData = state;
src.Value = char(newHeader);
headerCells = reshape(cellstr(state.editHeaderNames), 1, []);
tbl.ColumnName = [{'Variable', 'Value'}, headerCells];
end

function handleNamingModeChange(fig, ~, event)
state = fig.UserData;
state.nameMode = string(event.NewValue.Tag);
fig.UserData = state;
end

function layoutVehParamUI(fig)
if ~isvalid(fig)
    return;
end

state = fig.UserData;
if isempty(state) || ~isstruct(state) || ~isfield(state, 'ui') || ~isfield(state, 'numEditColumns')
    return;
end

u = state.ui;
n = state.numEditColumns;
figW = fig.Position(3);
figH = fig.Position(4);

left = 20;
right = 20;
gap = 8;
outputPanelY = 5;
outputPanelH = 90;
nameGroupY = outputPanelY + outputPanelH + 8;

filterLabelW = 110;
filterH = 22;
resetW = 130;
filterY = figH - 34;
filterX = left + filterLabelW + gap;
filterW = max(160, figW - right - resetW - gap - filterX);

u.filterLabel.Position = [left filterY filterLabelW filterH];
u.filterField.Position = [filterX filterY filterW filterH];
u.resetBtn.Position = [figW - right - resetW filterY resetW filterH];

tableBottom = 200;
tableTop = filterY - 36;
tableWidth = max(420, figW - left - right);
tableHeight = max(120, tableTop - tableBottom);
u.tbl.Position = [left tableBottom tableWidth tableHeight];

varW = 150;
valW = 150;
editAreaW = max(200, tableWidth - varW - valW);
baseEditW = floor(editAreaW / n);
editWidths = baseEditW * ones(1, n);
editWidths(end) = editWidths(end) + (editAreaW - baseEditW * n);
u.tbl.ColumnWidth = num2cell([varW, valW, editWidths]);

headerY = tableBottom + tableHeight + 4;
u.headerLabel.Position = [left headerY + 2 170 18];

headerX = left + varW + valW;
for idx = 1:n
    u.headerFields(idx).Position = [headerX headerY editWidths(idx) 22];
    headerX = headerX + editWidths(idx);
end

rowY = nameGroupY + 58;
u.newVarLabel.Position = [left rowY 85 22];

newVarX = left + 90;
newVarW = max(180, round(0.32 * tableWidth));
u.newVarField.Position = [newVarX rowY newVarW 22];

newValLabelX = newVarX + newVarW + 10;
u.newValLabel.Position = [newValLabelX rowY 85 22];

saveW = 160;
addRowW = 90;
saveX = figW - right - saveW;
addRowX = saveX - 14 - addRowW;
newValFieldX = newValLabelX + 85;
newValFieldW = max(120, addRowX - 10 - newValFieldX);
u.newValField.Position = [newValFieldX rowY newValFieldW 22];
u.addRowBtn.Position = [addRowX rowY 90 24];
u.saveBtn.Position = [saveX nameGroupY + 10 saveW 30];

nameGroupW = max(260, addRowX - 20 - left);
u.nameModeGroup.Position = [left nameGroupY nameGroupW 50];
u.outputPanel.Position = [left outputPanelY tableWidth outputPanelH];

panelW = u.outputPanel.Position(3);
u.outputLabel.Position = [10 34 max(160, panelW - 20) 22];
u.outputLink.Position = [10 6 max(160, panelW - 20) 18];
end

function updateOutputPanel(fig, folderPath)
state = fig.UserData;
state.outputFolderPath = string(folderPath);
fig.UserData = state;
updateOutputLink(fig);
end

function updateOutputLink(fig)
state = fig.UserData;
u = state.ui;
if strlength(state.outputFolderPath) == 0
    u.outputLink.Visible = 'off';
    u.outputLink.Tooltip = '';
    return;
end

u.outputLink.Text = "Open Folder";
u.outputLink.Tooltip = char(state.outputFolderPath);
u.outputLink.Visible = 'on';
end

function openSelectedOutputFolder(fig)
state = fig.UserData;
folderPath = char(state.outputFolderPath);
if isempty(folderPath)
    return;
end

if ~isfolder(folderPath)
    uialert(fig, sprintf('The folder no longer exists:\n%s', folderPath), 'Missing Folder');
    return;
end

try
    if ispc
        winopen(folderPath);
    elseif ismac
        system(['open "', folderPath, '" &']);
    else
        system(['xdg-open "', folderPath, '" >/dev/null 2>&1 &']);
    end
catch ME
    uialert(fig, sprintf('Unable to open folder:\n%s\n\n%s', folderPath, ME.message), 'Open Failed');
end
end

function addVariableRow(fig, tbl, filterField, newVarField, newValField, numEditColumns)
varName = strtrim(string(newVarField.Value));
defaultValue = strtrim(string(newValField.Value));

if strlength(varName) == 0
    uialert(fig, 'Enter a variable name before adding a row.', 'Missing Variable');
    return;
end

if strlength(defaultValue) == 0
    uialert(fig, 'Enter a default value before adding a row.', 'Missing Value');
    return;
end

state = fig.UserData;
existingNames = string(state.allTableData(:, 1));
if any(existingNames == varName)
    uialert(fig, sprintf('Variable "%s" already exists in the table.', char(varName)), 'Duplicate Variable');
    return;
end

newRow = [{char(varName), char(defaultValue)}, repmat({''}, 1, numEditColumns)];
state.allTableData(end + 1, :) = newRow;

newEntry = struct( ...
    'name', char(varName), ...
    'value', char(defaultValue), ...
    'valueStart', 0, ...
    'valueEnd', 0, ...
    'stmtEnd', 0);
state.entries(end + 1) = newEntry;

fig.UserData = state;
newVarField.Value = '';
newValField.Value = '';

applyVariableFilter(fig, tbl, filterField.Value);
end

function handleTableEdit(fig, event)
if isempty(event.Indices)
    return;
end

state = fig.UserData;
if isempty(state.visibleRowIdx)
    return;
end

visibleRow = event.Indices(1);
colIdx = event.Indices(2);
sourceRow = state.visibleRowIdx(visibleRow);
state.allTableData{sourceRow, colIdx} = event.NewData;
fig.UserData = state;
end

function saveChanges(fileText, selectedFolder, sourceMFileName, fig, numEditColumns)
state = fig.UserData;
editedData = state.allTableData;
entries = state.entries;
editHeaderNames = state.editHeaderNames;
namingMode = state.nameMode;
useHeaderNaming = (namingMode == "header");

sourceFolderName = string(getLeafName(selectedFolder));
timestamp = string(datetime('now', 'Format', 'yyyyMMdd_HHmmss'));
parentFolder = fileparts(selectedFolder);
hasAddedRows = any(arrayfun(@(e)e.valueStart == 0, entries));

createdFolders = strings(0, 1);
errors = strings(0, 1);

for editIndex = 1:numEditColumns
    currentColumnValues = editedData(:, 2 + editIndex);
    [columnEntries, hasAnyValue, updateNotes] = buildEntriesForEditColumn(entries, currentColumnValues);
    if ~hasAnyValue
        continue;
    end

    if useHeaderNaming
        variantTag = sanitizeFolderToken(editHeaderNames(editIndex), editIndex);
    else
        variantTag = "E" + string(editIndex);
    end
    folderBaseName = buildFolderBaseName(sourceFolderName, timestamp, variantTag, useHeaderNaming);

    [createdFolders, errors] = trackVariantCreation( ...
        createdFolders, errors, variantTag, fig, selectedFolder, parentFolder, ...
        folderBaseName, sourceFolderName, sourceMFileName, fileText, ...
        columnEntries, updateNotes, useHeaderNaming);
end

% If no edit columns are filled but new rows were added, create one baseline variant.
if isempty(createdFolders) && hasAddedRows
    [defaultEntries, defaultNotes] = buildDefaultAddedRowNotes(entries);
    if useHeaderNaming
        fallbackTag = "AddedDefaults";
    else
        fallbackTag = "E0";
    end
    folderBaseName = buildFolderBaseName(sourceFolderName, timestamp, fallbackTag, useHeaderNaming);
    [createdFolders, errors] = trackVariantCreation( ...
        createdFolders, errors, fallbackTag, fig, selectedFolder, parentFolder, ...
        folderBaseName, sourceFolderName, sourceMFileName, fileText, ...
        defaultEntries, defaultNotes, useHeaderNaming);
end

if isempty(createdFolders)
    uialert(fig, 'No Edit columns contain values and no new variables were added. No folders were created.', 'No Output');
    return;
end

updateOutputPanel(fig, parentFolder);

messageLines = "Created folder(s):" + newline + strjoin(createdFolders, newline);
if ~isempty(errors)
    messageLines = messageLines + newline + newline + "Warnings:" + newline + strjoin(errors, newline);
end
uialert(fig, char(messageLines), 'Done');

commitNote = sprintf( ...
    'Commit note: Base folder "%s"; auto-generated %d test folder(s): %s; updated %s and XML naming/content.', ...
    char(sourceFolderName), numel(createdFolders), getFolderListText(createdFolders), sourceMFileName);
fprintf('\n%s\n\n', commitNote);
end

function folderBaseName = buildFolderBaseName(sourceFolderName, timestamp, variantTag, useHeaderNaming)
if useHeaderNaming
    folderBaseName = sourceFolderName + "_" + variantTag;
else
    folderBaseName = sourceFolderName + "_" + timestamp + "_" + variantTag;
end
end

function [createdFolders, errors] = trackVariantCreation( ...
    createdFolders, errors, variantTag, fig, selectedFolder, parentFolder, ...
    folderBaseName, sourceFolderName, sourceMFileName, fileText, ...
    columnEntries, updateNotes, useHeaderNaming)
[isCreated, targetFolder, failureMessage] = createVariantFolder( ...
    fig, selectedFolder, parentFolder, folderBaseName, sourceFolderName, ...
    sourceMFileName, fileText, columnEntries, updateNotes, useHeaderNaming);
if isCreated
    createdFolders(end + 1) = string(targetFolder);
else
    errors(end + 1) = variantTag + ": " + failureMessage;
end
end

function [isCreated, targetFolder, failureMessage] = createVariantFolder( ...
    fig, selectedFolder, parentFolder, folderBaseName, sourceFolderName, ...
    sourceMFileName, fileText, columnEntries, updateNotes, useHeaderNaming)
isCreated = false;
failureMessage = "";

if useHeaderNaming
    targetFolder = fullfile(parentFolder, char(folderBaseName));
    if isfolder(targetFolder)
        userChoice = uiconfirm( ...
            fig, ...
            sprintf('Folder already exists:\n%s\n\nDo you want to replace it?', targetFolder), ...
            'Folder Exists', ...
            'Options', {'Replace', 'Cancel'}, ...
            'DefaultOption', 2, ...
            'CancelOption', 2, ...
            'Icon', 'warning');
        if strcmp(userChoice, 'Cancel')
            failureMessage = "operation cancelled by user";
            return;
        end

        try
            rmdir(targetFolder, 's');
        catch ME
            failureMessage = "failed to replace existing folder - " + string(ME.message);
            return;
        end
    end
else
    targetFolder = getUniqueFolderPath(parentFolder, folderBaseName);
end

[copied, copyMsg] = copyfile(selectedFolder, targetFolder);
if ~copied
    failureMessage = "copy failed - " + string(copyMsg);
    return;
end

targetMFilePath = fullfile(targetFolder, sourceMFileName);
try
    writeUpdatedVehParam(targetMFilePath, fileText, columnEntries, updateNotes);
catch ME
    failureMessage = string(sourceMFileName) + " update failed - " + string(ME.message);
    return;
end

try
    finalFolderName = string(getLeafName(targetFolder));
    updateFolderXml(targetFolder, char(sourceFolderName), char(finalFolderName));
catch ME
    failureMessage = "XML update failed - " + string(ME.message);
    return;
end

isCreated = true;
end

function safeToken = sanitizeFolderToken(rawName, fallbackIndex)
candidate = strtrim(string(rawName));
if strlength(candidate) == 0
    safeToken = "E" + string(fallbackIndex);
    return;
end

candidate = regexprep(candidate, '[<>:"/\\|?*]', '_');
candidate = regexprep(candidate, '\s+', '_');
candidate = regexprep(candidate, '_+', '_');
candidate = strip(candidate, '_');

if strlength(candidate) == 0
    safeToken = "E" + string(fallbackIndex);
else
    safeToken = candidate;
end
end

function folderListText = getFolderListText(createdFolders)
folderNames = strings(size(createdFolders));
for i = 1:numel(createdFolders)
    folderNames(i) = string(getLeafName(char(createdFolders(i))));
end
folderListText = char(strjoin(folderNames, ', '));
end

function [updatedEntries, hasAnyValue, updateNotes] = buildEntriesForEditColumn(entries, columnValues)
updatedEntries = entries;
entryCount = numel(entries);
pendingValues = strings(entryCount, 1);
changedMask = false(entryCount, 1);
addedMask = reshape(arrayfun(@(e)e.valueStart == 0, entries), [], 1);

for i = 1:entryCount
    candidateValue = strtrim(string(columnValues{i}));
    if strlength(candidateValue) > 0
        updatedEntries(i).value = char(candidateValue);
        pendingValues(i) = candidateValue;
        changedMask(i) = true;
    end
end

hasAnyValue = any(changedMask);
updateNotes = buildUpdateNotes(entries, pendingValues, changedMask, addedMask);
end

function [updatedEntries, updateNotes] = buildDefaultAddedRowNotes(entries)
updatedEntries = entries;
entryCount = numel(entries);
addedMask = reshape(arrayfun(@(e)e.valueStart == 0, entries), [], 1);
updateNotes = buildUpdateNotes(entries, strings(entryCount, 1), false(entryCount, 1), addedMask);
end

function updateNotes = buildUpdateNotes(entries, pendingValues, changedMask, addedMask)
noteIdx = find(changedMask | addedMask);
updateNotes = struct( ...
    'name', cell(numel(noteIdx), 1), ...
    'oldValue', cell(numel(noteIdx), 1), ...
    'newValue', cell(numel(noteIdx), 1), ...
    'isAdded', cell(numel(noteIdx), 1));

for k = 1:numel(noteIdx)
    idx = noteIdx(k);
    updateNotes(k).name = entries(idx).name;
    updateNotes(k).oldValue = char(entries(idx).value);
    if changedMask(idx)
        newValue = pendingValues(idx);
    else
        newValue = string(entries(idx).value);
    end
    updateNotes(k).newValue = char(newValue);
    updateNotes(k).isAdded = addedMask(idx);
end
end

function targetFolder = getUniqueFolderPath(parentFolder, folderBaseName)
targetFolder = fullfile(parentFolder, char(folderBaseName));
counter = 1;
while isfolder(targetFolder)
    targetFolder = fullfile(parentFolder, char(folderBaseName + "_" + string(counter)));
    counter = counter + 1;
end
end

function [entries, rawText] = parseVehParamFile(filePath)
rawText = fileread(filePath);

entries = struct('name', {}, 'value', {}, 'valueStart', {}, 'valueEnd', {}, 'stmtEnd', {});
ranges = getTopLevelStatementRanges(rawText);

for i = 1:size(ranges, 1)
    stmtStart = ranges(i, 1);
    stmtEnd = ranges(i, 2);
    statementText = rawText(stmtStart:stmtEnd);

    [localStart, trimmedStatement] = stripLeadingCommentLines(statementText);
    if isempty(trimmedStatement)
        continue;
    end

    [isAssignment, lhs, rhs, rhsStartRel, rhsEndRel] = parseAssignmentStatement(trimmedStatement);
    if ~isAssignment
        continue;
    end

    entries(end + 1).name = lhs; %#ok<AGROW>
    entries(end).value = rhs;
    entries(end).valueStart = stmtStart + localStart + rhsStartRel - 2;
    entries(end).valueEnd = stmtStart + localStart + rhsEndRel - 2;
    entries(end).stmtEnd = stmtEnd;
end
end

function ranges = getTopLevelStatementRanges(rawText)
ranges = zeros(0, 2);
if isempty(rawText)
    return;
end

i = 1;
n = length(rawText);
stmtStart = 1;
parenDepth = 0;
bracketDepth = 0;
braceDepth = 0;
inComment = false;
quoteMode = char(0);

while i <= n
    ch = rawText(i);

    if inComment
        if isLineBreak(ch)
            inComment = false;
        end
        i = i + 1;
        continue;
    end

    if quoteMode ~= char(0)
        if ch == quoteMode
            if quoteMode == '''' && i < n && rawText(i + 1) == ''''
                i = i + 2;
                continue;
            end
            quoteMode = char(0);
        end
        i = i + 1;
        continue;
    end

    if ch == '%'
        inComment = true;
    elseif ch == ''''
        if isLikelyStringDelimiter(rawText, i)
            quoteMode = '''';
        end
    elseif ch == '"'
        quoteMode = '"';
    elseif ch == '('
        parenDepth = parenDepth + 1;
    elseif ch == ')'
        parenDepth = max(parenDepth - 1, 0);
    elseif ch == '['
        bracketDepth = bracketDepth + 1;
    elseif ch == ']'
        bracketDepth = max(bracketDepth - 1, 0);
    elseif ch == '{'
        braceDepth = braceDepth + 1;
    elseif ch == '}'
        braceDepth = max(braceDepth - 1, 0);
    elseif ch == ';' && parenDepth == 0 && bracketDepth == 0 && braceDepth == 0
        if hasNonWhitespace(rawText(stmtStart:i))
            ranges(end + 1, :) = [stmtStart, i]; %#ok<AGROW>
        end
        stmtStart = i + 1;
    end

    i = i + 1;
end

if stmtStart <= n && hasNonWhitespace(rawText(stmtStart:n))
    ranges(end + 1, :) = [stmtStart, n];
end
end

function [localStart, trimmedStatement] = stripLeadingCommentLines(statementText)
n = length(statementText);
localStart = 1;

while localStart <= n
    while localStart <= n && isspace(statementText(localStart))
        localStart = localStart + 1;
    end
    if localStart > n
        trimmedStatement = '';
        return;
    end

    if statementText(localStart) ~= '%'
        break;
    end

    while localStart <= n && ~isLineBreak(statementText(localStart))
        localStart = localStart + 1;
    end
    while localStart <= n && isLineBreak(statementText(localStart))
        localStart = localStart + 1;
    end
end

trimmedStatement = statementText(localStart:end);
end

function [isAssignment, lhs, rhs, rhsStartRel, rhsEndRel] = parseAssignmentStatement(statementText)
isAssignment = false;
lhs = '';
rhs = '';
rhsStartRel = 0;
rhsEndRel = 0;

n = length(statementText);
if n == 0
    return;
end

eqIndex = findTopLevelAssignmentEquals(statementText);
if eqIndex == 0
    return;
end

lhsCandidate = strtrim(statementText(1:eqIndex - 1));
if isempty(lhsCandidate)
    return;
end

rhsLimit = n;
while rhsLimit >= 1 && isspace(statementText(rhsLimit))
    rhsLimit = rhsLimit - 1;
end
if rhsLimit >= 1 && statementText(rhsLimit) == ';'
    rhsLimit = rhsLimit - 1;
end
while rhsLimit >= 1 && isspace(statementText(rhsLimit))
    rhsLimit = rhsLimit - 1;
end

rhsStart = eqIndex + 1;
while rhsStart <= rhsLimit && isspace(statementText(rhsStart))
    rhsStart = rhsStart + 1;
end
if rhsStart > rhsLimit
    return;
end

lhs = lhsCandidate;
rhs = strtrim(statementText(rhsStart:rhsLimit));
rhsStartRel = rhsStart;
rhsEndRel = rhsLimit;
isAssignment = true;
end

function eqIndex = findTopLevelAssignmentEquals(text)
eqIndex = 0;
n = length(text);

i = 1;
parenDepth = 0;
bracketDepth = 0;
braceDepth = 0;
inComment = false;
quoteMode = char(0);

while i <= n
    ch = text(i);

    if inComment
        if isLineBreak(ch)
            inComment = false;
        end
        i = i + 1;
        continue;
    end

    if quoteMode ~= char(0)
        if ch == quoteMode
            if quoteMode == '''' && i < n && text(i + 1) == ''''
                i = i + 2;
                continue;
            end
            quoteMode = char(0);
        end
        i = i + 1;
        continue;
    end

    if ch == '%'
        inComment = true;
    elseif ch == ''''
        if isLikelyStringDelimiter(text, i)
            quoteMode = '''';
        end
    elseif ch == '"'
        quoteMode = '"';
    elseif ch == '('
        parenDepth = parenDepth + 1;
    elseif ch == ')'
        parenDepth = max(parenDepth - 1, 0);
    elseif ch == '['
        bracketDepth = bracketDepth + 1;
    elseif ch == ']'
        bracketDepth = max(bracketDepth - 1, 0);
    elseif ch == '{'
        braceDepth = braceDepth + 1;
    elseif ch == '}'
        braceDepth = max(braceDepth - 1, 0);
    elseif ch == '=' && parenDepth == 0 && bracketDepth == 0 && braceDepth == 0
        prevChar = previousNonWhitespace(text, i - 1);
        nextChar = nextNonWhitespace(text, i + 1);
        if any(prevChar == ['=', '<', '>', '~']) || nextChar == '='
            i = i + 1;
            continue;
        end
        eqIndex = i;
        return;
    end

    i = i + 1;
end
end

function out = previousNonWhitespace(text, startIndex)
out = char(0);
for i = startIndex:-1:1
    if ~isspace(text(i))
        out = text(i);
        return;
    end
end
end

function out = nextNonWhitespace(text, startIndex)
out = char(0);
n = length(text);
for i = startIndex:n
    if ~isspace(text(i))
        out = text(i);
        return;
    end
end
end

function tf = isLikelyStringDelimiter(text, quoteIndex)
prevChar = previousNonWhitespace(text, quoteIndex - 1);
if prevChar == char(0)
    tf = true;
    return;
end

delimiters = ['=', '(', '[', '{', ',', ';', ':', '+', '-', '*', '/', '\', '^', '~'];
tf = any(prevChar == delimiters);
end

function tf = hasNonWhitespace(text)
tf = ~isempty(strtrim(text));
end

function tf = isLineBreak(ch)
tf = (ch == newline);
end

function writeUpdatedVehParam(filePath, rawText, entries, updateNotes)
newContent = rawText;
if ~isempty(entries)
    hasExistingRange = arrayfun(@(e)e.valueStart > 0 && e.valueEnd >= e.valueStart, entries);
    existingEntries = entries(hasExistingRange);
    appendedEntries = entries(~hasExistingRange);

    if ~isempty(existingEntries)
        newContent = replaceExistingAssignments(rawText, existingEntries, appendedEntries);
    elseif ~isempty(appendedEntries)
        newContent = appendTextBlock(rawText, buildAppendedAssignments(appendedEntries));
    end
end

newContent = stripAutoUpdateSummary(newContent);
updateSummary = buildAutoUpdateSummary(updateNotes);
newContent = [updateSummary, newline, newline, newContent];

fid = fopen(filePath, 'w');
if fid < 0
    error('Unable to open Veh_Param.m for writing.');
end
cleanupObj = onCleanup(@() fclose(fid));
fwrite(fid, newContent, 'char');
end

function newContent = replaceExistingAssignments(rawText, existingEntries, appendedEntries)
nChars = length(rawText);
[~, sortOrder] = sort([existingEntries.valueStart]);
existingEntries = existingEntries(sortOrder);
insertAfterPos = max([existingEntries.stmtEnd]);

segments = cell(1, 2 * numel(existingEntries) + 3);
segIdx = 1;
cursor = 1;

for i = 1:numel(existingEntries)
    valueStart = existingEntries(i).valueStart;
    valueEnd = existingEntries(i).valueEnd;
    if valueStart < cursor || valueEnd < valueStart
        error('Invalid assignment range detected while writing Veh_Param.m.');
    end

    segments{segIdx} = rawText(cursor:valueStart - 1);
    segIdx = segIdx + 1;
    segments{segIdx} = char(existingEntries(i).value);
    segIdx = segIdx + 1;
    cursor = valueEnd + 1;
end

if isempty(appendedEntries)
    segments{segIdx} = rawText(cursor:end);
    newContent = [segments{1:segIdx}];
    return;
end

splitPos = min(max(insertAfterPos, cursor - 1), nChars);
if splitPos >= cursor
    segments{segIdx} = rawText(cursor:splitPos);
else
    segments{segIdx} = '';
end
segIdx = segIdx + 1;

segments{segIdx} = [newline, newline, buildAppendedAssignments(appendedEntries)];
segIdx = segIdx + 1;

if splitPos < nChars
    segments{segIdx} = rawText(splitPos + 1:end);
else
    segments{segIdx} = '';
end
newContent = [segments{1:segIdx}];
end

function outText = appendTextBlock(inText, blockText)
if isempty(strtrim(inText))
    outText = blockText;
    return;
end

if inText(end) ~= newline
    inText = [inText, newline];
end
outText = [inText, newline, blockText];
end

function appendBlock = buildAppendedAssignments(appendedEntries)
lines = strings(numel(appendedEntries) + 2, 1);
lines(1) = "% Added variables (from GUI Add Row)";
for i = 1:numel(appendedEntries)
    lines(i + 1) = string(appendedEntries(i).name) + " = " + string(appendedEntries(i).value) + ";";
end
lines(end) = "% End added variables";
appendBlock = char(strjoin(lines, newline));
end

function summaryText = buildAutoUpdateSummary(updateNotes)
isAddedMask = reshape(logical([updateNotes.isAdded]), [], 1);
summaryLines = ["% Auto-update summary (CreateDIVeDatasetVariant)"; ...
    buildSummarySection("% Updated variables (default -> updated):", updateNotes(~isAddedMask)); ...
    buildSummarySection("% Added variables (default -> final):", updateNotes(isAddedMask)); ...
    "% End auto-update summary"];
summaryText = char(strjoin(summaryLines, newline));
end

function sectionLines = buildSummarySection(titleLine, notes)
if isempty(notes)
    sectionLines = [titleLine; "%   (none)"];
    return;
end

sectionLines = strings(numel(notes) + 1, 1);
sectionLines(1) = titleLine;
for k = 1:numel(notes)
    previousValue = sanitizeSummaryValue(notes(k).oldValue);
    currentValue = sanitizeSummaryValue(notes(k).newValue);
    sectionLines(k + 1) = "%   " + string(notes(k).name) + ": " + previousValue + " -> " + currentValue;
end
end

function cleanValue = sanitizeSummaryValue(rawValue)
cleanValue = regexprep(string(rawValue), '\s+', ' ');
end

function outText = stripAutoUpdateSummary(inText)
summaryPattern = [ ...
    '^% Auto-update summary \(CreateDIVeDatasetVariant\)(?:\r\n|\n|\r)', ...
    '(?:%.*(?:\r\n|\n|\r))*?', ...
    '% End auto-update summary(?:\r\n|\n|\r)*'];
outText = regexprep(inText, summaryPattern, '', 'once');
end

function updateFolderXml(folderPath, oldFolderName, newFolderName)
xmlFiles = dir(fullfile(folderPath, '*.xml'));
if isempty(xmlFiles)
    error('No XML file found in duplicated folder: %s', folderPath);
end

preferredXml = fullfile(folderPath, [oldFolderName, '.xml']);
if isfile(preferredXml)
    sourceXml = preferredXml;
else
    sourceXml = fullfile(folderPath, xmlFiles(1).name);
end

targetXml = fullfile(folderPath, [newFolderName, '.xml']);
if ~strcmpi(sourceXml, targetXml)
    movefile(sourceXml, targetXml, 'f');
end

xmlText = fileread(targetXml);

% Replace only exact quoted folder-name tokens: "Old" -> "New".
if ~isempty(oldFolderName) && ~strcmp(oldFolderName, newFolderName)
    quotedOld = ['"', oldFolderName, '"'];
    quotedNew = ['"', newFolderName, '"'];
    xmlText = strrep(xmlText, quotedOld, quotedNew);
end

fid = fopen(targetXml, 'w');
if fid < 0
    error('Unable to write XML file: %s', targetXml);
end
cleanupObj = onCleanup(@() fclose(fid));
fwrite(fid, xmlText, 'char');
end

function leaf = getLeafName(folderPath)
[~, leaf] = fileparts(folderPath);
end
