% eBus_Release_Check.m
% Compares two eBus release folders containing matching MAT files,
% exports a comparative KPI workbook, saves default figures, and
% generates template-based reports for each release.

clearvars;
clc;

selectedFolders = selectTwoReleaseFolders(pwd);
if isempty(selectedFolders)
    fprintf('Folder selection cancelled. Script stopped.\n');
    return;
end

[pairInfos, relativeMatPaths, comparisonInfo] = collectComparableMatPairs(selectedFolders{1}, selectedFolders{2});
if isempty(pairInfos)
    fprintf('No MAT files found in the selected release folders.\n');
    return;
end

releaseLabels = buildReleaseLabels(selectedFolders);
resultsBaseDir = getResultsBaseDir(selectedFolders);

thisScriptDir = fileparts(mfilename('fullpath'));
kpiBankPath = fullfile(thisScriptDir, '..', 'KPIs_Plots', 'eBus_KPIs_Plots_Bank.xlsx');
templateDir = fullfile(thisScriptDir, '..', 'Report_Templates', 'Batch_Sim_Report_Templates');

if ~isfile(kpiBankPath)
    error('KPI bank not found at: %s', kpiBankPath);
end

fprintf('Reading KPI and plot configuration...\n');
kpiConfig = readtable(kpiBankPath, ...
    'Sheet', 'KPIs', ...
    'VariableNamingRule', 'preserve', ...
    'TextType', 'string');

plotConfig = readtable(kpiBankPath, ...
    'Sheet', 'Plots', ...
    'VariableNamingRule', 'preserve', ...
    'TextType', 'string');

plotProperties = readtable(kpiBankPath, ...
    'Sheet', 'Plot Properties', ...
    'VariableNamingRule', 'preserve', ...
    'TextType', 'string');
colorMap = buildColorMap(plotProperties);

requiredCols = ["Group", "KPI", "Unit", "Variable Name", "Print"];
for iCol = 1:numel(requiredCols)
    if ~ismember(requiredCols(iCol), string(kpiConfig.Properties.VariableNames))
        error('Required column "%s" was not found in sheet "KPIs".', requiredCols(iCol));
    end
end

eqCols = string(kpiConfig.Properties.VariableNames);
eqCols = eqCols(startsWith(eqCols, "KPI Equation"));
if isempty(eqCols)
    error('No equation columns (KPI Equation 1..N) found in sheet "KPIs".');
end

groupCol = string(kpiConfig.("Group"));
kpiCol = string(kpiConfig.("KPI"));
validRows = strlength(strtrim(groupCol)) > 0 & strlength(strtrim(kpiCol)) > 0;
kpiConfig = kpiConfig(validRows, :);

numKpis = height(kpiConfig);
printProcessedMatFileInfo(pairInfos, releaseLabels, selectedFolders, comparisonInfo);

fprintf('Creating results folder...\n');
resultsDir = createDiveResultsFolder(resultsBaseDir);
groups = string(kpiConfig.("Group"));
kpis = string(kpiConfig.("KPI"));
variableNames = string(kpiConfig.("Variable Name"));
units = string(kpiConfig.("Unit"));
if ismember("Sr No", string(kpiConfig.Properties.VariableNames))
    srNos = kpiConfig.("Sr No");
else
    srNos = (1:numKpis).';
end

numPairs = numel(pairInfos);
pairResults = repmat(struct( ...
    'RelativeMatPath', "", ...
    'MatFileName', "", ...
    'ReleaseLabelA', "", ...
    'ReleaseLabelB', "", ...
    'SourcePathA', "", ...
    'SourcePathB', "", ...
    'PairResultsDir', "", ...
    'ResultMatrix', strings(numKpis, 2), ...
    'ResultValueMatrix', strings(numKpis, 2)), numPairs, 1);
generatedReportCount = 0;

fprintf('Computing comparative KPIs for %d MAT file entry(ies)...\n', numPairs);
for iPair = 1:numPairs
    fprintf('  [%d/%d] %s\n', iPair, numPairs, char(relativeMatPaths(iPair)));
    hasA = strlength(string(pairInfos(iPair).SourcePathA)) > 0 && isfile(char(string(pairInfos(iPair).SourcePathA)));
    hasB = strlength(string(pairInfos(iPair).SourcePathB)) > 0 && isfile(char(string(pairInfos(iPair).SourcePathB)));

    contextA = struct();
    contextB = struct();
    resultMatrixA = strings(numKpis, 1);
    resultValueMatrixA = strings(numKpis, 1);
    resultMatrixB = strings(numKpis, 1);
    resultValueMatrixB = strings(numKpis, 1);

    if hasA
        contextA = load(pairInfos(iPair).SourcePathA);
        [resultMatrixA, resultValueMatrixA, contextA] = computeKpisForContext(kpiConfig, eqCols, contextA);
    end
    if hasB
        contextB = load(pairInfos(iPair).SourcePathB);
        [resultMatrixB, resultValueMatrixB, contextB] = computeKpisForContext(kpiConfig, eqCols, contextB);
    end

    pairResultMatrix = [resultMatrixA, resultMatrixB];
    pairValueMatrix = [resultValueMatrixA, resultValueMatrixB];
    pairDir = createPairResultsFolder(resultsDir, iPair, relativeMatPaths(iPair));

    activeLabels = strings(1, 2);
    activeContexts = cell(2, 1);
    activeResultMatrix = strings(numKpis, 2);
    activeCount = 0;
    if hasA
        activeCount = activeCount + 1;
        activeLabels(activeCount) = releaseLabels(1);
        activeContexts{activeCount, 1} = contextA;
        activeResultMatrix(:, activeCount) = resultMatrixA;
    end
    if hasB
        activeCount = activeCount + 1;
        activeLabels(activeCount) = releaseLabels(2);
        activeContexts{activeCount, 1} = contextB;
        activeResultMatrix(:, activeCount) = resultMatrixB;
    end

    if activeCount > 0
        activeLabels = activeLabels(1:activeCount);
        activeContexts = activeContexts(1:activeCount);
        activeResultMatrix = activeResultMatrix(:, 1:activeCount);
        fprintf('      Saving default figure files...\n');
        figSaveDirs = createPerMatFigureGroupFolders(pairDir, activeLabels, "Default");
        plotForAllFiles(plotConfig, activeContexts, activeLabels, colorMap, "", false, figSaveDirs);

        fprintf('      Generating reports from batch templates...\n');
        generatedReportCount = generatedReportCount + generateReportsForAllTemplates( ...
            templateDir, pairDir, activeLabels, variableNames, activeResultMatrix, ...
            plotConfig, activeContexts, colorMap);
    end

    pairResults(iPair).RelativeMatPath = string(relativeMatPaths(iPair));
    pairResults(iPair).MatFileName = string(pairInfos(iPair).MatFileName);
    pairResults(iPair).ReleaseLabelA = string(releaseLabels(1));
    pairResults(iPair).ReleaseLabelB = string(releaseLabels(2));
    pairResults(iPair).SourcePathA = string(pairInfos(iPair).SourcePathA);
    pairResults(iPair).SourcePathB = string(pairInfos(iPair).SourcePathB);
    pairResults(iPair).PairResultsDir = string(pairDir);
    pairResults(iPair).ResultMatrix = pairResultMatrix;
    pairResults(iPair).ResultValueMatrix = pairValueMatrix;
end

summaryCells = buildReleaseComparisonSummaryCells(pairResults, groups, kpis, units, srNos);
summaryTable = cell2table(summaryCells(2:end, :), ...
    'VariableNames', buildSummaryTableVariableNames(summaryCells(1, :)));

fprintf('Writing comparative KPI Excel report...\n');
exportReleaseComparisonExcel(resultsDir, pairResults, groups, kpis, units, srNos, releaseLabels, kpiBankPath);

assignin('base', 'eBusReleaseComparisonSummaryTable', summaryTable);
assignin('base', 'eBusReleaseComparisonResultsDir', resultsDir);

fprintf('\neBus release comparison complete.\n');
if generatedReportCount > 0
    fprintf('Comparative Excel report, default figure files, and template reports were created.\n');
else
    fprintf('Comparative Excel report and default figure files were created.\n');
end
printResultFolderLink(resultsDir);
try
    winopen(resultsDir);
catch
end

function printProcessedMatFileInfo(pairInfos, releaseLabels, selectedFolders, comparisonInfo)
nFiles = numel(pairInfos);
fprintf('\nMAT files prepared for comparison (%d):\n', nFiles);
fprintf('  Release 1 Folder Path: %s\n', char(string(selectedFolders{1})));
fprintf('  Release 2 Folder Path: %s\n', char(string(selectedFolders{2})));

matchedMask = false(nFiles, 1);
release1Names = strings(nFiles, 1);
release2Names = strings(nFiles, 1);
for iFile = 1:nFiles
    release1Names(iFile) = getFileNameOnly(pairInfos(iFile).SourcePathA);
    release2Names(iFile) = getFileNameOnly(pairInfos(iFile).SourcePathB);
    matchedMask(iFile) = strlength(release1Names(iFile)) > 0 && strlength(release2Names(iFile)) > 0;
end

fprintf('\n');
printTwoColumnNameTable('Matched MAT Files', releaseLabels, release1Names(matchedMask), release2Names(matchedMask));

if (isfield(comparisonInfo, 'OnlyInA') && ~isempty(comparisonInfo.OnlyInA)) || ...
        (isfield(comparisonInfo, 'OnlyInB') && ~isempty(comparisonInfo.OnlyInB))
    onlyInA = getMatNamesFromRelativePaths(comparisonInfo.OnlyInA);
    onlyInB = getMatNamesFromRelativePaths(comparisonInfo.OnlyInB);
    fprintf('\n');
    printTwoColumnNameTable('Non-Identical MAT Files', releaseLabels, onlyInA, onlyInB);
end

fprintf('\n');
end

function printTwoColumnNameTable(titleText, releaseLabels, leftValues, rightValues)
leftValues = string(leftValues(:));
rightValues = string(rightValues(:));
rowCount = max(numel(leftValues), numel(rightValues));
if rowCount == 0
    return;
end

if numel(leftValues) < rowCount
    leftValues(end + 1:rowCount, 1) = "";
end
if numel(rightValues) < rowCount
    rightValues(end + 1:rowCount, 1) = "";
end

col1Header = string(releaseLabels(1));
col2Header = string(releaseLabels(2));
colGap = '    ';
col1Width = max([strlength(col1Header); strlength(leftValues)]);
col2Width = max([strlength(col2Header); strlength(rightValues)]);

fprintf('%s:\n', char(string(titleText)));
headerLine = padRightLocal(col1Header, col1Width) + string(colGap) + padRightLocal(col2Header, col2Width);
separatorLine = string(repmat('_', 1, col1Width)) + string(colGap) + string(repmat('_', 1, col2Width));

fprintf('%s\n', char(headerLine));
fprintf('%s\n', char(separatorLine));
for iRow = 1:rowCount
    rowLine = padRightLocal(leftValues(iRow), col1Width) + string(colGap) + padRightLocal(rightValues(iRow), col2Width);
    fprintf('%s\n', char(rowLine));
end
end

function fileName = getFileNameOnly(filePath)
fileName = "";
if strlength(string(filePath)) == 0
    return;
end
[~, baseName, ext] = fileparts(char(string(filePath)));
fileName = string(baseName) + string(ext);
end

function fileNames = getMatNamesFromRelativePaths(relativePaths)
relativePaths = string(relativePaths(:));
fileNames = strings(numel(relativePaths), 1);
for iPath = 1:numel(relativePaths)
    [~, baseName, ext] = fileparts(char(relativePaths(iPath)));
    fileNames(iPath) = string(baseName) + string(ext);
end
end

function out = padRightLocal(inText, totalWidth)
inText = string(inText);
padCount = max(0, totalWidth - strlength(inText));
out = inText + string(repmat(' ', 1, padCount));
end

function [resultMatrix, resultValueMatrix, context] = computeKpisForContext(kpiConfig, eqCols, context)
numKpis = height(kpiConfig);
resultMatrix = strings(numKpis, 1);
resultValueMatrix = strings(numKpis, 1);

for iKpi = 1:numKpis
    solved = false;
    solvedValue = [];

    for iEq = 1:numel(eqCols)
        eqText = string(kpiConfig{iKpi, eqCols(iEq)});
        eqText = strtrim(eqText);
        if strlength(eqText) == 0 || ismissing(eqText)
            continue;
        end

        [candidateValue, ok] = tryEvaluateEquation(eqText, context);
        if ~ok || ~isSupportedScalar(candidateValue)
            continue;
        end

        solved = true;
        solvedValue = candidateValue;
        break;
    end

    unitText = string(kpiConfig.("Unit")(iKpi));
    if solved
        resultMatrix(iKpi) = formatValueWithUnit(solvedValue, unitText);
        resultValueMatrix(iKpi) = formatValueOnly(solvedValue);
        varName = strtrim(string(kpiConfig.("Variable Name")(iKpi)));
        if strlength(varName) > 0 && isvarname(varName)
            context.(varName) = solvedValue;
        end
    else
        resultMatrix(iKpi) = "N.A";
        resultValueMatrix(iKpi) = "N.A";
    end
end
end

function [absDiffText, pctDiffText] = buildDiffTexts(valueA, valueB)
[numA, okA] = tryParseNumericValue(valueA);
[numB, okB] = tryParseNumericValue(valueB);
if ~(okA && okB)
    absDiffText = "N.A";
    pctDiffText = "N.A";
    return;
end

absDiff = abs(numB - numA);
absDiffText = formatValueOnly(absDiff);
if abs(numA) > eps
    pctDiffText = formatValueOnly(absDiff / abs(numA) * 100);
elseif absDiff <= eps
    pctDiffText = "0";
else
    pctDiffText = "N.A";
end
end

function generatedCount = generateReportsForAllTemplates(templateDir, resultsDir, fileLabels, variableNames, resultMatrix, ...
    plotConfig, contextsByFile, colorMap)
generatedCount = 0;

[templateNames, templatePaths] = getTemplateDocuments(templateDir);
if isempty(templateNames)
    fprintf('No batch report templates found.\n');
    return;
end

reportDirs = createPerMatReportFolders(resultsDir, fileLabels);
numTemplates = numel(templateNames);
numFiles = numel(fileLabels);

for iTemplate = 1:numTemplates
    templateName = templateNames(iTemplate);
    templatePath = templatePaths{iTemplate};
    fprintf('  %s %s\n', char(9670), char(templateName));

    for iFile = 1:numFiles
        outputPath = buildReportOutputPath(reportDirs{iFile}, templateName, fileLabels(iFile));
        copyfile(templatePath, outputPath, 'f');

        placeholderMap = buildPlaceholderMapForFile(variableNames, resultMatrix(:, iFile));
        replacePlaceholdersInDocument(outputPath, placeholderMap);
        replaceFigurePlaceholdersInReport(outputPath, plotConfig, contextsByFile{iFile}, fileLabels(iFile), colorMap);
        generatedCount = generatedCount + 1;
    end
end
end

function [templateNames, templatePaths] = getTemplateDocuments(templateDir)
templateNames = strings(0, 1);
templatePaths = {};
if ~isfolder(templateDir)
    return;
end

items = dir(templateDir);
items = items(~[items.isdir]);
if isempty(items)
    return;
end

templateNames = string({items.name})';
templatePaths = cell(numel(items), 1);
for i = 1:numel(items)
    templatePaths{i} = fullfile(items(i).folder, items(i).name);
end
end

function selectedFolders = selectTwoReleaseFolders(startDir)
selectedFolders = {};
if nargin < 1 || strlength(string(startDir)) == 0 || ~isfolder(startDir)
    startDir = pwd;
end

folder1Path = "";
folder2Path = "";
dialogSize = [100 100 760 220];

try
    dlg = dialog( ...
        'Name', 'Select Release Folders', ...
        'Position', dialogSize, ...
        'WindowStyle', 'modal', ...
        'Resize', 'off');
catch
    dlg = [];
end

if isempty(dlg) || ~isgraphics(dlg)
    selectedFolders = fallbackSelectTwoFolders(startDir);
    return;
end

setappdata(dlg, 'cancelled', true);
set(dlg, 'CloseRequestFcn', @onCancel);

uicontrol(dlg, 'Style', 'text', ...
    'Position', [35 150 110 22], ...
    'String', 'Release Folder 1', ...
    'HorizontalAlignment', 'left', ...
    'FontWeight', 'bold');
edit1 = uicontrol(dlg, 'Style', 'edit', ...
    'Position', [145 145 480 32], ...
    'HorizontalAlignment', 'left', ...
    'BackgroundColor', 'white', ...
    'Enable', 'inactive', ...
    'String', '');
uicontrol(dlg, 'Style', 'pushbutton', ...
    'Position', [640 145 90 32], ...
    'String', 'Browse...', ...
    'Callback', @(~, ~)browseForFolder(1));

uicontrol(dlg, 'Style', 'text', ...
    'Position', [35 95 110 22], ...
    'String', 'Release Folder 2', ...
    'HorizontalAlignment', 'left', ...
    'FontWeight', 'bold');
edit2 = uicontrol(dlg, 'Style', 'edit', ...
    'Position', [145 90 480 32], ...
    'HorizontalAlignment', 'left', ...
    'BackgroundColor', 'white', ...
    'Enable', 'inactive', ...
    'String', '');
uicontrol(dlg, 'Style', 'pushbutton', ...
    'Position', [640 90 90 32], ...
    'String', 'Browse...', ...
    'Callback', @(~, ~)browseForFolder(2));

uicontrol(dlg, 'Style', 'pushbutton', ...
    'Position', [510 25 100 36], ...
    'String', 'Continue', ...
    'FontWeight', 'bold', ...
    'Callback', @onContinue);
uicontrol(dlg, 'Style', 'pushbutton', ...
    'Position', [625 25 100 36], ...
    'String', 'Cancel', ...
    'Callback', @onCancel);

uiwait(dlg);

if ~isgraphics(dlg)
    return;
end

cancelled = getappdata(dlg, 'cancelled');
if ~cancelled
    selectedFolders = {char(folder1Path); char(folder2Path)};
end

if isgraphics(dlg)
    delete(dlg);
end

selectedFolders = unique(selectedFolders, 'stable');
if numel(selectedFolders) ~= 2
    selectedFolders = {};
end

    function browseForFolder(folderIndex)
        currentStart = startDir;
        if folderIndex == 1 && strlength(folder1Path) > 0 && isfolder(char(folder1Path))
            currentStart = char(folder1Path);
        elseif folderIndex == 2 && strlength(folder2Path) > 0 && isfolder(char(folder2Path))
            currentStart = char(folder2Path);
        elseif folderIndex == 2 && strlength(folder1Path) > 0 && isfolder(char(folder1Path))
            currentStart = char(folder1Path);
        end

        selectedPath = uigetdir(currentStart, sprintf('Select Release Folder %d', folderIndex));
        if isequal(selectedPath, 0)
            return;
        end

        if folderIndex == 1
            folder1Path = string(selectedPath);
            set(edit1, 'String', char(folder1Path));
        else
            folder2Path = string(selectedPath);
            set(edit2, 'String', char(folder2Path));
        end
    end

    function onContinue(~, ~)
        if strlength(folder1Path) == 0 || ~isfolder(char(folder1Path))
            errordlg('Please choose Folder 1.', 'Missing Folder', 'modal');
            return;
        end
        if strlength(folder2Path) == 0 || ~isfolder(char(folder2Path))
            errordlg('Please choose Folder 2.', 'Missing Folder', 'modal');
            return;
        end
        if strcmpi(char(folder1Path), char(folder2Path))
            errordlg('Folder 1 and Folder 2 must be different.', 'Invalid Selection', 'modal');
            return;
        end

        setappdata(dlg, 'cancelled', false);
        uiresume(dlg);
    end

    function onCancel(~, ~)
        if isgraphics(dlg)
            setappdata(dlg, 'cancelled', true);
            uiresume(dlg);
        end
    end
end

function selectedFolders = fallbackSelectTwoFolders(startDir)
selectedFolders = {};
promptTitles = {'Select Release Folder 1', 'Select Release Folder 2'};
tempFolders = cell(2, 1);
for i = 1:2
    selectedPath = uigetdir(char(string(startDir)), promptTitles{i});
    if isequal(selectedPath, 0)
        return;
    end
    tempFolders{i} = char(string(selectedPath));
    startDir = tempFolders{i};
end
selectedFolders = unique(tempFolders, 'stable');
end

function releaseLabels = buildReleaseLabels(selectedFolders)
releaseLabels = strings(1, numel(selectedFolders));
for iFolder = 1:numel(selectedFolders)
    [~, folderName] = fileparts(char(string(selectedFolders{iFolder})));
    if strlength(string(folderName)) == 0
        folderName = sprintf('Release_%02d', iFolder);
    end
    releaseLabels(iFolder) = string(folderName);
end
releaseLabels = string(matlab.lang.makeUniqueStrings(cellstr(releaseLabels)));
end

function [pairInfos, relativeMatPaths, comparisonInfo] = collectComparableMatPairs(folderA, folderB)
pairInfos = repmat(struct( ...
    'RelativeMatPath', "", ...
    'MatFileName', "", ...
    'SourcePathA', "", ...
    'SourcePathB', ""), 0, 1);
relativeMatPaths = strings(0, 1);
comparisonInfo = struct('MatchedCount', 0, 'OnlyInA', strings(0, 1), 'OnlyInB', strings(0, 1));

itemsA = getMatFilesRecursive(folderA);
itemsB = getMatFilesRecursive(folderB);
if isempty(itemsA) && isempty(itemsB)
    return;
end

[fullPathsA, relPathsA] = collectMatPathsWithRelativeNames(folderA, itemsA);
[fullPathsB, relPathsB] = collectMatPathsWithRelativeNames(folderB, itemsB);

normRelA = normalizeRelativePaths(relPathsA);
normRelB = normalizeRelativePaths(relPathsB);

[unionNormRel, ~] = sort(unique([normRelA; normRelB], 'stable'));
if isempty(unionNormRel)
    return;
end

relativeMatPaths = strings(numel(unionNormRel), 1);
pairInfos = repmat(struct( ...
    'RelativeMatPath', "", ...
    'MatFileName', "", ...
    'SourcePathA', "", ...
    'SourcePathB', ""), numel(unionNormRel), 1);

comparisonInfo.OnlyInA = relPathsA(~ismember(normRelA, normRelB));
comparisonInfo.OnlyInB = relPathsB(~ismember(normRelB, normRelA));

for iPair = 1:numel(unionNormRel)
    thisNormPath = unionNormRel(iPair);
    idxA = find(normRelA == thisNormPath, 1, 'first');
    idxB = find(normRelB == thisNormPath, 1, 'first');

    displayRelPath = "";
    if ~isempty(idxA)
        displayRelPath = relPathsA(idxA);
    elseif ~isempty(idxB)
        displayRelPath = relPathsB(idxB);
    end
    relativeMatPaths(iPair) = displayRelPath;

    [~, baseName, ext] = fileparts(char(displayRelPath));
    pairInfos(iPair).RelativeMatPath = string(displayRelPath);
    pairInfos(iPair).MatFileName = string([baseName ext]);
    if ~isempty(idxA)
        pairInfos(iPair).SourcePathA = string(fullPathsA{idxA});
    else
        pairInfos(iPair).SourcePathA = "";
    end
    if ~isempty(idxB)
        pairInfos(iPair).SourcePathB = string(fullPathsB{idxB});
    else
        pairInfos(iPair).SourcePathB = "";
    end
    if ~isempty(idxA) && ~isempty(idxB)
        comparisonInfo.MatchedCount = comparisonInfo.MatchedCount + 1;
    end
end
end

function [fullPaths, relPaths] = collectMatPathsWithRelativeNames(rootDir, items)
fullPaths = {};
relPaths = strings(0, 1);
if isempty(items)
    return;
end

fullPaths = cell(numel(items), 1);
relPaths = strings(numel(items), 1);
for iItem = 1:numel(items)
    fullPaths{iItem} = fullfile(items(iItem).folder, items(iItem).name);
    relPaths(iItem) = makeRelativePath(rootDir, fullPaths{iItem});
end
end

function normPaths = normalizeRelativePaths(relPaths)
normPaths = lower(string(relPaths));
normPaths = replace(normPaths, '/', '\');
end

function msg = composeMatMismatchMessage(folderA, folderB, missingInA, missingInB)
lines = strings(0, 1);
lines(end + 1, 1) = "Selected folders do not contain identical MAT file sets.";
lines(end + 1, 1) = "Folder 1: " + string(folderA);
lines(end + 1, 1) = "Folder 2: " + string(folderB);

if ~isempty(missingInB)
    lines(end + 1, 1) = "";
    lines(end + 1, 1) = "Present in Folder 1 but missing in Folder 2:";
    lines = [lines; reshape(formatPreviewList(missingInB, 10), [], 1)];
end
if ~isempty(missingInA)
    lines(end + 1, 1) = "";
    lines(end + 1, 1) = "Present in Folder 2 but missing in Folder 1:";
    lines = [lines; reshape(formatPreviewList(missingInA, 10), [], 1)];
end

msg = strjoin(lines, newline);
end

function lines = formatPreviewList(values, maxItems)
values = string(values(:));
values = values(strlength(values) > 0);
if nargin < 2
    maxItems = 10;
end
showCount = min(numel(values), maxItems);
lines = reshape("  - " + values(1:showCount), [], 1);
if numel(values) > showCount
    lines(end + 1, 1) = sprintf('  ... and %d more', numel(values) - showCount);
end
end

function selectedFolders = selectMatFolders(startDir)
selectedFolders = {};
if nargin < 1 || strlength(string(startDir)) == 0 || ~isfolder(startDir)
    startDir = pwd;
end

if usejava('awt')
    try
        chooser = javax.swing.JFileChooser(java.io.File(char(string(startDir))));
        chooser.setDialogTitle('Select one or more folders containing MAT files');
        chooser.setFileSelectionMode(javax.swing.JFileChooser.DIRECTORIES_ONLY);
        chooser.setMultiSelectionEnabled(true);

        status = chooser.showOpenDialog([]);
        if status == javax.swing.JFileChooser.APPROVE_OPTION
            javaFiles = chooser.getSelectedFiles();
            selectedFolders = cell(numel(javaFiles), 1);
            for i = 1:numel(javaFiles)
                selectedFolders{i} = char(javaFiles(i).getAbsolutePath());
            end
        end
    catch
        selectedFolders = {};
    end
end

if isempty(selectedFolders)
    selectedPath = uigetdir(char(string(startDir)), 'Select folder containing MAT files');
    if isequal(selectedPath, 0)
        return;
    end
    selectedFolders = {char(string(selectedPath))};
end

selectedFolders = unique(selectedFolders, 'stable');
end

function [matFilePaths, relativeMatPaths] = collectMatFilesFromFolders(selectedFolders)
matFilePaths = {};
relativeMatPaths = strings(0, 1);
if isempty(selectedFolders)
    return;
end

for iFolder = 1:numel(selectedFolders)
    rootDir = char(string(selectedFolders{iFolder}));
    matItems = getMatFilesRecursive(rootDir);
    if isempty(matItems)
        continue;
    end

    [~, rootLabel] = fileparts(rootDir);
    if strlength(string(rootLabel)) == 0
        rootLabel = string(rootDir);
    else
        rootLabel = string(rootLabel);
    end

    for iFile = 1:numel(matItems)
        fullPath = fullfile(matItems(iFile).folder, matItems(iFile).name);
        relPath = makeRelativePath(rootDir, fullPath);
        displayPath = rootLabel;
        if strlength(relPath) > 0
            displayPath = string(fullfile(char(rootLabel), char(relPath)));
        end

        matFilePaths{end + 1, 1} = fullPath; %#ok<AGROW>
        relativeMatPaths(end + 1, 1) = string(displayPath); %#ok<AGROW>
    end
end

if isempty(matFilePaths)
    return;
end

[uniquePaths, uniqueIdx] = unique(lower(string(matFilePaths)), 'stable');
matFilePaths = matFilePaths(uniqueIdx);
relativeMatPaths = relativeMatPaths(uniqueIdx);

[~, sortIdx] = sort(uniquePaths);
matFilePaths = matFilePaths(sortIdx);
relativeMatPaths = relativeMatPaths(sortIdx);
end

function resultsBaseDir = getResultsBaseDir(selectedFolders)
if isempty(selectedFolders)
    resultsBaseDir = pwd;
    return;
end

if isscalar(selectedFolders)
    resultsBaseDir = char(string(selectedFolders{1}));
    return;
end

resultsBaseDir = findCommonParentFolder(selectedFolders);
if strlength(string(resultsBaseDir)) == 0 || ~isfolder(resultsBaseDir)
    resultsBaseDir = pwd;
end
end

function matItems = getMatFilesRecursive(rootDir)
matItems = dir(fullfile(rootDir, '**', '*.mat'));
if isempty(matItems)
    matItems = dir(fullfile(rootDir, '*.mat'));
end
if isempty(matItems)
    return;
end
matItems = matItems(~[matItems.isdir]);
end

function relPath = makeRelativePath(rootDir, filePath)
rootDir = char(string(rootDir));
filePath = char(string(filePath));

if startsWith(filePath, [rootDir filesep], 'IgnoreCase', ispc)
    relPath = string(filePath(numel(rootDir) + 2:end));
elseif strcmpi(filePath, rootDir)
    relPath = string(filePath);
else
    relPath = string(filePath);
end
end

function commonDir = findCommonParentFolder(paths)
commonDir = char(string(paths{1}));
for i = 2:numel(paths)
    thisPath = char(string(paths{i}));
    while ~isempty(commonDir)
        if strcmpi(thisPath, commonDir) || startsWith(thisPath, [commonDir filesep], 'IgnoreCase', ispc)
            break;
        end
        nextDir = fileparts(commonDir);
        if strcmp(nextDir, commonDir)
            commonDir = '';
            break;
        end
        commonDir = nextDir;
    end

    if isempty(commonDir)
        break;
    end
end
end

function fileLabel = buildMatFileLabel(relativeMatPath)
relativeMatPath = char(string(relativeMatPath));
[folderPart, baseName, ~] = fileparts(relativeMatPath);
if isempty(folderPart)
    fileLabel = string(baseName);
    return;
end

folderPart = strrep(folderPart, '/', ' | ');
folderPart = strrep(folderPart, '\', ' | ');
fileLabel = string([folderPart ' | ' baseName]);
end

function printResultFolderLink(resultsDir)
if nargin < 1 || strlength(string(resultsDir)) == 0 || ~isfolder(resultsDir)
    return;
end

[~, folderName] = fileparts(char(string(resultsDir)));
folderCmd = sprintf('matlab:winopen(''%s'')', escapeForMatlabCharLiteral(resultsDir));
fprintf('\n');
fprintf('Result Folder: <a href="%s">%s</a>\n', folderCmd, folderName);
fprintf('\n');
end

function out = escapeForMatlabCharLiteral(inText)
out = strrep(char(string(inText)), '''', '''''');
end

function outDir = createDiveResultsFolder(baseDir)
timeStamp = char(datetime('now', 'Format', 'yyyyMMdd_HHmmss'));
folderName = [timeStamp '_eBus_Release_Results'];
outDir = fullfile(baseDir, folderName);

if ~isfolder(outDir)
    mkdir(outDir);
    return;
end

suffix = 1;
while isfolder(outDir)
    outDir = fullfile(baseDir, sprintf('%s_%02d', folderName, suffix));
    suffix = suffix + 1;
end
mkdir(outDir);
end

function figDir = createFigureOutputFolder(resultsDir, folderLabel)
figDir = "";
if strlength(string(resultsDir)) == 0 || ~isfolder(resultsDir)
    return;
end

if nargin < 2 || strlength(strtrim(string(folderLabel))) == 0
    folderLabel = "Default";
end

safeLabel = string(sanitizeFileName(folderLabel));
folderName = char(safeLabel + "_Fig");
figDir = fullfile(char(string(resultsDir)), folderName);
if ~isfolder(figDir)
    mkdir(figDir);
end
end

function [matDirs, figBaseDirs, reportBaseDirs] = ensurePerMatOutputFolders(resultsDir, fileLabels)
nFiles = numel(fileLabels);
matDirs = repmat({''}, nFiles, 1);
figBaseDirs = repmat({''}, nFiles, 1);
reportBaseDirs = repmat({''}, nFiles, 1);
if strlength(string(resultsDir)) == 0 || ~isfolder(resultsDir)
    return;
end

safeLabels = strings(nFiles, 1);
for iFile = 1:nFiles
    safeLabels(iFile) = string(sanitizeFileName(fileLabels(iFile)));
    if strlength(strtrim(safeLabels(iFile))) == 0
        safeLabels(iFile) = sprintf('MatFile_%02d', iFile);
    end
end
safeLabels = string(matlab.lang.makeUniqueStrings(cellstr(safeLabels)));

for iFile = 1:nFiles
    matDir = fullfile(char(string(resultsDir)), char(safeLabels(iFile)));
    if ~isfolder(matDir)
        mkdir(matDir);
    end

    figDir = fullfile(matDir, 'Fig');
    if ~isfolder(figDir)
        mkdir(figDir);
    end

    reportDir = fullfile(matDir, 'Report');
    if ~isfolder(reportDir)
        mkdir(reportDir);
    end

    matDirs{iFile} = matDir;
    figBaseDirs{iFile} = figDir;
    reportBaseDirs{iFile} = reportDir;
end
end

function figSaveDirs = createPerMatFigureGroupFolders(resultsDir, fileLabels, groupLabel)
[~, figBaseDirs, ~] = ensurePerMatOutputFolders(resultsDir, fileLabels);
nFiles = numel(figBaseDirs);
figSaveDirs = repmat({''}, nFiles, 1);
for iFile = 1:nFiles
    baseDir = string(figBaseDirs{iFile});
    if strlength(baseDir) == 0 || ~isfolder(char(baseDir))
        continue;
    end
    figSaveDirs{iFile} = char(createFigureOutputFolder(char(baseDir), groupLabel));
end
end

function reportDirs = createPerMatReportFolders(resultsDir, fileLabels)
[~, ~, reportBaseDirs] = ensurePerMatOutputFolders(resultsDir, fileLabels);
reportDirs = reportBaseDirs;
end

function safeName = sanitizeFileName(inName)
safeName = regexprep(char(string(inName)), '[^\w\-\.\(\) ]', '_');
safeName = strtrim(safeName);
if isempty(safeName)
    safeName = 'Figure';
end
end

function outPath = makeUniqueFilePath(folderPath, baseName, ext)
outPath = fullfile(folderPath, [baseName ext]);
if ~isfile(outPath)
    return;
end

suffix = 1;
while true
    candidate = fullfile(folderPath, sprintf('%s_%03d%s', baseName, suffix, ext));
    if ~isfile(candidate)
        outPath = candidate;
        return;
    end
    suffix = suffix + 1;
end
end

function pairDir = createPairResultsFolder(resultsDir, pairIndex, relativeMatPath)
pairLabel = sanitizeFileName(buildMatFileLabel(relativeMatPath));
if strlength(string(pairLabel)) > 60
    pairLabel = extractBefore(string(pairLabel), 61);
end
folderName = sprintf('%03d_%s', pairIndex, char(string(pairLabel)));
pairDir = fullfile(resultsDir, folderName);
if ~isfolder(pairDir)
    mkdir(pairDir);
end
end

function exportReleaseComparisonExcel(resultsDir, pairResults, groups, kpis, units, srNos, releaseLabels, kpiBankPath)
if ~isfolder(resultsDir)
    return;
end

[~, folderName] = fileparts(resultsDir);
outPath = fullfile(resultsDir, [char(folderName) '.xlsx']);

if nargin >= 8 && strlength(string(kpiBankPath)) > 0 && isfile(char(string(kpiBankPath)))
    copyfile(char(string(kpiBankPath)), outPath, 'f');
end

summaryCells = buildReleaseComparisonSummaryCells(pairResults, groups, kpis, units, srNos);
writeComparisonSheet(outPath, 'KPIs', summaryCells);

matchedFileCells = buildMatchedFileCells(pairResults, releaseLabels);
writecell(matchedFileCells, outPath, 'Sheet', 'Matched_Files', 'Range', 'A1');

for iPair = 1:numel(pairResults)
    sheetName = buildComparisonSheetName(iPair, pairResults(iPair).MatFileName);
    pairCells = buildReleaseComparisonPairCells(pairResults(iPair), groups, kpis, units, srNos, releaseLabels);
    writecell(pairCells, outPath, 'Sheet', sheetName, 'Range', 'A1');
end

postProcessComparisonWorkbook(outPath, size(summaryCells, 1), size(summaryCells, 2), numel(pairResults));
end

function writeComparisonSheet(workbookPath, sheetName, cells)
clearRows = size(cells, 1);
clearCols = size(cells, 2);
if isfile(workbookPath)
    try
        baseCells = readcell(workbookPath, 'Sheet', sheetName);
        if ~isempty(baseCells)
            [baseRows, baseCols] = size(baseCells);
            clearRows = max(clearRows, baseRows);
            clearCols = max(clearCols, baseCols);
        end
    catch
    end
end

outCells = repmat({''}, clearRows, clearCols);
outCells(1:size(cells, 1), 1:size(cells, 2)) = cells;
writecell(outCells, workbookPath, 'Sheet', sheetName, 'Range', 'A1');
end

function cells = buildReleaseComparisonSummaryCells(pairResults, groups, kpis, units, srNos)
numPairs = numel(pairResults);
numKpis = numel(kpis);
cells = repmat({''}, numKpis + 1, 4 + 4 * numPairs);
cells(1, 1:4) = {'Sr No', 'Group', 'KPI', 'Unit'};

for iKpi = 1:numKpis
    rowIdx = iKpi + 1;
    cells{rowIdx, 1} = toExcelCellValue(srNos(iKpi), '');
    cells{rowIdx, 2} = toExcelCellValue(groups(iKpi), '');
    cells{rowIdx, 3} = toExcelCellValue(kpis(iKpi), '');
    cells{rowIdx, 4} = toExcelCellValue(units(iKpi), '');
end

for iPair = 1:numPairs
    colStart = 5 + (iPair - 1) * 4;
    pairLabel = buildPairSummaryLabel(pairResults(iPair).MatFileName);
    cells{1, colStart} = sprintf('%s_%s', char(string(pairResults(iPair).ReleaseLabelA)), pairLabel);
    cells{1, colStart + 1} = sprintf('%s_%s', char(string(pairResults(iPair).ReleaseLabelB)), pairLabel);
    cells{1, colStart + 2} = 'Absolute Difference';
    cells{1, colStart + 3} = '% Difference';

    for iKpi = 1:numKpis
        rowIdx = iKpi + 1;
        [absDiff, pctDiff] = buildDiffExcelValues(pairResults(iPair).ResultValueMatrix(iKpi, 1), ...
            pairResults(iPair).ResultValueMatrix(iKpi, 2));
        cells{rowIdx, colStart} = toExcelCellValue(pairResults(iPair).ResultValueMatrix(iKpi, 1), '');
        cells{rowIdx, colStart + 1} = toExcelCellValue(pairResults(iPair).ResultValueMatrix(iKpi, 2), '');
        cells{rowIdx, colStart + 2} = absDiff;
        cells{rowIdx, colStart + 3} = pctDiff;
    end
end
end

function label = buildPairSummaryLabel(matFileName)
[~, baseName, ~] = fileparts(char(string(matFileName)));
label = regexprep(baseName, '[^\w]', '_');
if isempty(label)
    label = 'MAT';
end
end

function variableNames = buildSummaryTableVariableNames(headerCells)
headerStrings = string(headerCells);
headerStrings(ismissing(headerStrings) | strlength(headerStrings) == 0) = "Column";
variableNames = matlab.lang.makeValidName(cellstr(headerStrings));
variableNames = matlab.lang.makeUniqueStrings(variableNames);
end

function cells = buildMatchedFileCells(pairResults, releaseLabels)
numPairs = numel(pairResults);
cells = repmat({''}, numPairs + 1, 5);
cells(1, :) = {'MAT File', char(string(releaseLabels(1))), char(string(releaseLabels(2))), 'Relative MAT Path', 'Pair Result Folder'};
for iPair = 1:numPairs
    cells{iPair + 1, 1} = char(string(pairResults(iPair).MatFileName));
    cells{iPair + 1, 2} = char(string(pairResults(iPair).SourcePathA));
    cells{iPair + 1, 3} = char(string(pairResults(iPair).SourcePathB));
    cells{iPair + 1, 4} = char(string(pairResults(iPair).RelativeMatPath));
    cells{iPair + 1, 5} = char(string(pairResults(iPair).PairResultsDir));
end
end

function cells = buildReleaseComparisonPairCells(pairResult, groups, kpis, units, srNos, releaseLabels)
numKpis = numel(kpis);
cells = repmat({''}, numKpis + 4, 9);
cells(1, 1:2) = {'MAT File', char(string(pairResult.MatFileName))};
cells(2, 1:2) = {'Relative MAT Path', char(string(pairResult.RelativeMatPath))};
cells(3, 1:2) = {char(string(releaseLabels(1))), char(string(pairResult.SourcePathA))};
cells(3, 3:4) = {char(string(releaseLabels(2))), char(string(pairResult.SourcePathB))};
cells(4, :) = {'Sr No', 'Group', 'KPI', 'Unit', char(string(releaseLabels(1))), ...
    char(string(releaseLabels(2))), 'Absolute Difference', ...
    sprintf('%% Difference vs %s', char(string(releaseLabels(1)))), 'Pair Result Folder'};

for iKpi = 1:numKpis
    rowIdx = iKpi + 4;
    [absDiff, pctDiff] = buildDiffExcelValues(pairResult.ResultValueMatrix(iKpi, 1), pairResult.ResultValueMatrix(iKpi, 2));
    cells{rowIdx, 1} = toExcelCellValue(srNos(iKpi), '');
    cells{rowIdx, 2} = toExcelCellValue(groups(iKpi), '');
    cells{rowIdx, 3} = toExcelCellValue(kpis(iKpi), '');
    cells{rowIdx, 4} = toExcelCellValue(units(iKpi), '');
    cells{rowIdx, 5} = toExcelCellValue(pairResult.ResultValueMatrix(iKpi, 1), '');
    cells{rowIdx, 6} = toExcelCellValue(pairResult.ResultValueMatrix(iKpi, 2), '');
    cells{rowIdx, 7} = absDiff;
    cells{rowIdx, 8} = pctDiff;
    cells{rowIdx, 9} = char(string(pairResult.PairResultsDir));
end
end

function [absDiffValue, pctDiffValue] = buildDiffExcelValues(valueA, valueB)
[isMissingA, isMissingB] = deal(isUnavailableComparisonValue(valueA), isUnavailableComparisonValue(valueB));
if isMissingA || isMissingB
    absDiffValue = '';
    pctDiffValue = '';
    return;
end

[numA, okA] = tryParseNumericValue(valueA);
[numB, okB] = tryParseNumericValue(valueB);
if ~(okA && okB)
    absDiffValue = 'N.A';
    pctDiffValue = 'N.A';
    return;
end

absDiffValue = round(abs(numB - numA), 2);
if abs(numA) > eps
    pctDiffValue = round(abs(numB - numA) / abs(numA) * 100, 2);
elseif absDiffValue <= eps
    pctDiffValue = 0;
else
    pctDiffValue = 'N.A';
end
end

function tf = isUnavailableComparisonValue(valueText)
txt = strtrim(string(valueText));
tf = ismissing(txt) || strlength(txt) == 0;
end

function sheetName = buildComparisonSheetName(pairIndex, matFileName)
baseName = sprintf('Pair_%03d_%s', pairIndex, char(string(matFileName)));
sheetName = regexprep(baseName, '[:\\/\?\*\[\]]', '_');
if strlength(string(sheetName)) > 31
    sheetName = extractBefore(string(sheetName), 32);
end
sheetName = char(string(sheetName));
end

function postProcessComparisonWorkbook(workbookPath, summaryRows, summaryCols, numPairs)
if ~ispc || ~isfile(workbookPath)
    return;
end

excelApp = [];
wb = [];
try
    excelApp = actxserver('Excel.Application');
    excelApp.DisplayAlerts = false;
    excelApp.Visible = false;
    wb = excelApp.Workbooks.Open(workbookPath, false, false);
catch
    cleanupExcelSession(excelApp, wb);
    return;
end

try
    sheetNamesToDelete = ["Plots", "Plot Properties"];
    for i = 1:numel(sheetNamesToDelete)
        if wb.Worksheets.Count <= 1
            break;
        end
        try
            wb.Worksheets.Item(char(sheetNamesToDelete(i))).Delete;
        catch
        end
    end

    for iSheet = 1:double(wb.Worksheets.Count)
        ws = wb.Worksheets.Item(iSheet);
        try
            loCount = double(ws.ListObjects.Count);
            for iLo = loCount:-1:1
                ws.ListObjects.Item(iLo).Unlist;
            end
        catch
        end
        try
            ws.Cells.UnMerge;
        catch
        end
        try
            ws.Cells.FormatConditions.Delete;
        catch
        end
        try
            usedRange = ws.UsedRange;
            usedRange.Font.Name = 'Consolas';
            usedRange.Font.Size = 10;
            usedRange.WrapText = false;
            usedRange.HorizontalAlignment = -4108;
            usedRange.VerticalAlignment = -4108;
            usedRange.Columns.AutoFit;
        catch
        end
    end

    try
        wsKpi = wb.Worksheets.Item('KPIs');
        rngAll = wsKpi.Range(xlA1(1, 1), xlA1(summaryRows, summaryCols));
        rngAll.Font.Name = 'Consolas';
        rngAll.Font.Size = 11;
        rngAll.WrapText = false;
        rngAll.HorizontalAlignment = -4108;
        rngAll.VerticalAlignment = -4108;
        wsKpi.Rows.Item(1).Font.Bold = true;

        for iPair = 1:numPairs
            diffCol = 5 + (iPair - 1) * 4 + 2;
            pctCol = diffCol + 1;
            diffRange = wsKpi.Range(xlA1(2, diffCol), xlA1(summaryRows, diffCol));
            pctRange = wsKpi.Range(xlA1(2, pctCol), xlA1(summaryRows, pctCol));
            diffRange.NumberFormat = '0.00';
            pctRange.NumberFormat = '0.00';
        end
        wsKpi.Columns.AutoFit;
    catch
    end

    wb.Save;
catch
end

cleanupExcelSession(excelApp, wb);
end

function exportKpiResultsExcel(resultsDir, data)
if ~isfolder(resultsDir)
    return;
end

[~, folderName] = fileparts(resultsDir);
fileName = [char(folderName) '.xlsx'];
outPath = fullfile(resultsDir, fileName);

sourcePath = "";
if isfield(data, 'kpiBankPath')
    sourcePath = string(data.kpiBankPath);
end
if strlength(sourcePath) > 0 && isfile(sourcePath)
    copyfile(char(sourcePath), outPath, 'f');
end

writeKpiSheetFallback(outPath, data);
postProcessExportedWorkbook(outPath, data);
end

function writeKpiSheetFallback(outPath, data)
nRows = numel(data.kpis);
nFiles = numel(data.fileLabels);
addDiffCols = (nFiles == 2);
extraCols = 2 * double(addDiffCols);
desiredRows = nRows + 1;
desiredCols = 4 + nFiles + extraCols;

clearRows = desiredRows;
clearCols = desiredCols;
if isfile(outPath)
    try
        baseCells = readcell(outPath, 'Sheet', 'KPIs');
        if ~isempty(baseCells)
            [baseRows, baseCols] = size(baseCells);
            clearRows = max(clearRows, baseRows);
            clearCols = max(clearCols, baseCols);
        end
    catch
    end
end

headers = cell(1, desiredCols);
headers(1:4) = {'Sr No', 'Group', 'KPI', 'Unit'};
for iFile = 1:nFiles
    headers{4 + iFile} = char(string(data.fileLabels(iFile)));
end
if addDiffCols
    headers{4 + nFiles + 1} = sprintf('Difference (%s - %s)', ...
        char(string(data.fileLabels(1))), char(string(data.fileLabels(2))));
    headers{4 + nFiles + 2} = 'Difference %';
end

cells = repmat({''}, clearRows, clearCols);
cells(1, 1:desiredCols) = headers;
for iRow = 1:nRows
    cells{iRow + 1, 1} = toExcelCellValue(data.srNos(iRow), '');
    cells{iRow + 1, 2} = toExcelCellValue(data.groups(iRow), '');
    cells{iRow + 1, 3} = toExcelCellValue(data.kpis(iRow), '');
    cells{iRow + 1, 4} = toExcelCellValue(data.units(iRow), '');

    for iFile = 1:nFiles
        cells{iRow + 1, 4 + iFile} = toExcelCellValue(data.resultValueMatrix(iRow, iFile), 'N.A');
    end

    if addDiffCols
        [v1, ok1] = tryParseNumericValue(data.resultValueMatrix(iRow, 1));
        [v2, ok2] = tryParseNumericValue(data.resultValueMatrix(iRow, 2));
        diffCol = 4 + nFiles + 1;
        pctCol = 4 + nFiles + 2;
        if ok1 && ok2
            diffVal = round(v1 - v2, 2);
            cells{iRow + 1, diffCol} = diffVal;
            if abs(v2) > eps
                cells{iRow + 1, pctCol} = round((v1 - v2) / v2 * 100, 2);
            else
                cells{iRow + 1, pctCol} = 'N.A';
            end
        else
            cells{iRow + 1, diffCol} = 'N.A';
            cells{iRow + 1, pctCol} = 'N.A';
        end
    end
end

writecell(cells, outPath, 'Sheet', 'KPIs', 'Range', 'A1');
end

function postProcessExportedWorkbook(workbookPath, data)
if ~ispc || ~isfile(workbookPath)
    return;
end

excelApp = [];
wb = [];
try
    excelApp = actxserver('Excel.Application');
    excelApp.DisplayAlerts = false;
    excelApp.Visible = false;
    wb = excelApp.Workbooks.Open(workbookPath, false, false);
catch
    cleanupExcelSession(excelApp, wb);
    return;
end

sheetNames = ["Plots", "Plot Properties"];
for i = 1:numel(sheetNames)
    if wb.Worksheets.Count <= 1
        break;
    end
    sName = char(string(sheetNames(i)));
    try
        ws = wb.Worksheets.Item(sName);
        ws.Delete;
    catch
    end
end

wsKpi = [];
try
    wsKpi = wb.Worksheets.Item('KPIs');
catch
    try
        wsKpi = wb.Worksheets.Item(1);
    catch
    end
end

if ~isempty(wsKpi)
    nRows = numel(data.kpis);
    nFiles = numel(data.fileLabels);
    hasDiffCols = (nFiles == 2);
    lastCol = 4 + nFiles + 2 * double(hasDiffCols);
    lastRow = max(1, nRows + 1);

    try
        loCount = double(wsKpi.ListObjects.Count);
        for iLo = loCount:-1:1
            wsKpi.ListObjects.Item(iLo).Unlist;
        end
    catch
    end
    try
        wsKpi.Cells.UnMerge;
    catch
    end
    try
        wsKpi.Cells.FormatConditions.Delete;
    catch
    end

    try
        rngAll = wsKpi.Range(xlA1(1, 1), xlA1(lastRow, lastCol));
        rngAll.Font.Name = 'Consolas';
        rngAll.Font.Size = 11;
        rngAll.Font.Bold = false;
        rngAll.HorizontalAlignment = -4108;
        rngAll.VerticalAlignment = -4108;
        rngAll.WrapText = false;
    catch
    end

    if hasDiffCols && nRows > 0
        diffCol = 4 + nFiles + 1;
        pctCol = 4 + nFiles + 2;
        try
            diffRange = wsKpi.Range(xlA1(2, diffCol), xlA1(nRows + 1, diffCol));
            pctRange = wsKpi.Range(xlA1(2, pctCol), xlA1(nRows + 1, pctCol));
            diffRange.NumberFormat = '0.00';
            pctRange.NumberFormat = '0.00';
            diffRange.HorizontalAlignment = -4108;
            pctRange.HorizontalAlignment = -4108;
            diffRange.VerticalAlignment = -4108;
            pctRange.VerticalAlignment = -4108;
        catch
        end

        lightRed = 13551615;
        for iRow = 1:nRows
            [v1, ok1] = tryParseNumericValue(data.resultValueMatrix(iRow, 1));
            [v2, ok2] = tryParseNumericValue(data.resultValueMatrix(iRow, 2));
            if ok1 && ok2
                diffVal = round(v1 - v2, 2);
                if abs(diffVal) > eps
                    try
                        wsKpi.Range(xlA1(iRow + 1, diffCol)).Interior.Color = lightRed;
                        wsKpi.Range(xlA1(iRow + 1, pctCol)).Interior.Color = lightRed;
                    catch
                    end
                end
            end
        end
    end
end

try
    wb.Save;
catch
end

cleanupExcelSession(excelApp, wb);
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

function [val, ok] = tryParseNumericValue(inVal)
ok = false;
val = NaN;

if isnumeric(inVal) && isscalar(inVal) && ~isnan(inVal)
    val = double(inVal);
    ok = true;
    return;
end

s = strtrim(string(inVal));
if ismissing(s) || strlength(s) == 0
    return;
end

[numVal, isNum] = parseNumericLiteral(char(s));
if isNum
    val = numVal;
    ok = true;
end
end

function out = toExcelCellValue(inVal, defaultText)
if nargin < 2
    defaultText = '';
end

if iscell(inVal) && isscalar(inVal)
    inVal = inVal{1};
end

if isempty(inVal)
    out = char(string(defaultText));
    return;
end

if isstring(inVal)
    if isscalar(inVal)
        if ismissing(inVal)
            out = char(string(defaultText));
            return;
        end
        txt = strtrim(char(inVal));
    else
        inVal = inVal(~ismissing(inVal));
        if isempty(inVal)
            out = char(string(defaultText));
            return;
        end
        txt = strtrim(char(strjoin(inVal, " ")));
    end

    if isempty(txt)
        out = char(string(defaultText));
        return;
    end

    [numVal, isNum] = parseNumericLiteral(txt);
    if isNum
        out = numVal;
    else
        out = txt;
    end
    return;
end

if ischar(inVal)
    txt = strtrim(inVal);
    if isempty(txt)
        out = char(string(defaultText));
        return;
    end

    [numVal, isNum] = parseNumericLiteral(txt);
    if isNum
        out = numVal;
    else
        out = txt;
    end
    return;
end

if isnumeric(inVal) || islogical(inVal)
    if isscalar(inVal) && ~(isnumeric(inVal) && isnan(inVal))
        out = inVal;
    else
        out = char(string(defaultText));
    end
    return;
end

out = char(string(defaultText));
end

function [numVal, isNum] = parseNumericLiteral(txt)
txt = strtrim(char(txt));
numVal = NaN;
isNum = false;
if isempty(txt) || strcmpi(txt, 'N.A') || strcmpi(txt, 'NA')
    return;
end

pattern = '^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$';
if isempty(regexp(txt, pattern, 'once'))
    return;
end

numCandidate = str2double(txt);
if ~isnan(numCandidate)
    numVal = numCandidate;
    isNum = true;
end
end

function outPath = buildReportOutputPath(targetDir, templateName, fileLabel)
[~, baseName, ext] = fileparts(char(templateName));
safeReportName = regexprep(baseName, '[^\w\- ]', '_');
safeMatName = regexprep(char(string(fileLabel)), '[^\w\- ]', '_');
timeStampPrefix = char(datetime('now', 'Format', 'yyyyMMdd_HHmmss'));

% Naming convention: <YYYYMMDD_HHMMSS>_<matfilename>_<SelectedReportName>
candidate = fullfile(targetDir, [timeStampPrefix '_' safeMatName '_' safeReportName ext]);

if ~isfile(candidate)
    outPath = candidate;
    return;
end

suffix = char(datetime('now', 'Format', 'HHmmssSSS'));
outPath = fullfile(targetDir, [timeStampPrefix '_' safeMatName '_' safeReportName '_' suffix ext]);
end

function placeholderMap = buildPlaceholderMapForFile(variableNames, fileValues)
placeholderMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
for i = 1:numel(variableNames)
    varName = strtrim(string(variableNames(i)));
    if strlength(varName) == 0 || ismissing(varName)
        continue;
    end
    if ~isvarname(varName)
        continue;
    end
    key = ['*' char(varName)];
    valueStr = strtrim(string(fileValues(i)));
    if ismissing(valueStr) || strlength(valueStr) == 0
        valueStr = "N.A [NA]";
    elseif strcmpi(valueStr, "N.A") || strcmpi(valueStr, "NA")
        valueStr = "N.A [NA]";
    end
    value = char(valueStr);
    placeholderMap(key) = value;
end
end

function replacePlaceholdersInDocument(docPath, placeholderMap)
[~, ~, ext] = fileparts(docPath);
ext = lower(string(ext));
openXmlExt = [".pptx",".pptm",".docx",".docm",".xlsx",".xlsm"];
plainTextExt = [".txt",".md",".csv",".json",".xml",".html",".htm"];

if (ext == ".docx" || ext == ".docm" || ext == ".doc") && ispc
    okWord = replaceTextPlaceholdersInWordViaCom(docPath, placeholderMap);
    if okWord
        return;
    end
end

if any(ext == openXmlExt)
    replaceInOpenXmlPackage(docPath, placeholderMap);
elseif any(ext == plainTextExt)
    replaceInTextFile(docPath, placeholderMap);
else
    % Unknown format: keep copied template as-is.
end
end

function replaceFigurePlaceholdersInReport(docPath, plotConfig, context, fileLabel, colorMap)
[~, ~, ext] = fileparts(docPath);
ext = lower(string(ext));
if ext == ".docx" || ext == ".docm" || ext == ".doc"
    replaceFigurePlaceholdersInWord(docPath, plotConfig, context, fileLabel, colorMap);
    return;
end
if ext ~= ".pptx" && ext ~= ".pptm"
    return;
end
if ~hasFigurePlaceholdersInPpt(docPath)
    return;
end
replaceFigurePlaceholdersInPpt(docPath, plotConfig, context, fileLabel, colorMap);
end

function replaceInTextFile(filePath, placeholderMap)
content = fileread(filePath);
keys = placeholderMap.keys;
for i = 1:numel(keys)
    key = keys{i};
    content = strrep(content, key, placeholderMap(key));
end
content = annotateUnresolvedPlaceholdersInText(content);
writeTextFile(filePath, content);
end

function replaceInOpenXmlPackage(filePath, placeholderMap)
workDir = tempname;
mkdir(workDir);
cleanupObj = onCleanup(@()cleanupTempDir(workDir));

unzip(filePath, workDir);
xmlFiles = dir(fullfile(workDir, '**', '*.xml'));
if isempty(xmlFiles)
    return;
end

keys = placeholderMap.keys;
xmlSafeValues = containers.Map('KeyType', 'char', 'ValueType', 'char');
for i = 1:numel(keys)
    k = keys{i};
    xmlSafeValues(k) = escapeXmlText(placeholderMap(k));
end

for i = 1:numel(xmlFiles)
    xmlPath = fullfile(xmlFiles(i).folder, xmlFiles(i).name);
    content = fileread(xmlPath);
    updated = content;
    for j = 1:numel(keys)
        key = keys{j};
        updated = replaceXmlPlaceholderTokens(updated, key, xmlSafeValues(key));
    end
    updated = annotateUnresolvedPlaceholdersInXml(updated);
    if ~strcmp(updated, content)
        writeTextFile(xmlPath, updated);
    end
end

zipPath = [tempname '.zip'];
rootItems = dir(workDir);
rootItems = rootItems(~ismember({rootItems.name}, {'.', '..'}));
if isempty(rootItems)
    return;
end
rootNames = {rootItems.name};
zip(zipPath, rootNames, workDir);
movefile(zipPath, filePath, 'f');
end

function out = escapeXmlText(inText)
out = inText;
out = strrep(out, '&', '&amp;');
out = strrep(out, '<', '&lt;');
out = strrep(out, '>', '&gt;');
out = strrep(out, '"', '&quot;');
out = strrep(out, '''', '&apos;');
end

function contentOut = replaceXmlPlaceholderTokens(contentIn, key, valueXmlSafe)
% Replace direct token occurrences first.
contentOut = strrep(contentIn, key, valueXmlSafe);

% Also replace placeholders split across multiple text runs (common in PPTX/DOCX),
% e.g. <a:t>*</a:t> ... <a:t>veh_weight</a:t>.
contentOut = replaceSplitTokenForPrefix(contentOut, key, valueXmlSafe, 'a');
contentOut = replaceSplitTokenForPrefix(contentOut, key, valueXmlSafe, 'w');
end

function contentOut = replaceSplitTokenForPrefix(contentIn, key, valueXmlSafe, prefix)
contentOut = contentIn;
if isempty(key)
    return;
end

tOpenPattern = ['<' prefix ':t(?:\s+[^>]*)?>'];
tClosePattern = ['</' prefix ':t>'];
tCloseTag = ['</' prefix ':t>'];
rPrPattern = ['(?:<' prefix ':rPr[^>]*/>\s*|<' prefix ':rPr[^>]*>(?:\s*<[^>]+>\s*)*</' prefix ':rPr>\s*)?'];

boundary = [tClosePattern ...
    '\s*</' prefix ':r>\s*<' prefix ':r[^>]*>\s*' ...
    rPrPattern ...
    tOpenPattern];

tokenPattern = regexptranslate('escape', key(1));
for i = 2:numel(key)
    tokenPattern = [tokenPattern '(?:' boundary ')?' regexptranslate('escape', key(i))]; %#ok<AGROW>
end

fullPattern = ['(' tOpenPattern ')' tokenPattern tClosePattern];
replacement = ['$1' valueXmlSafe tCloseTag];
contentOut = regexprep(contentOut, fullPattern, replacement);
end

function contentOut = annotateUnresolvedPlaceholdersInText(contentIn)
% Add [NA] for any unresolved placeholder token like *veh_weight.
% Exclude **Figure[...] placeholders from generic KPI placeholder annotation.
pattern = '(?<![\w\*])(\*(?![Ff]igure\b)[A-Za-z][A-Za-z0-9_]*)(?!\s*\[NA\])';
contentOut = regexprep(contentIn, pattern, '$1 [NA]');
end

function contentOut = annotateUnresolvedPlaceholdersInXml(contentIn)
% First annotate placeholders that are in a single text run.
contentOut = annotateUnresolvedPlaceholdersInText(contentIn);

% Then annotate placeholders split across runs (common in PPTX/DOCX).
contentOut = annotateSplitUnresolvedForPrefix(contentOut, 'a');
contentOut = annotateSplitUnresolvedForPrefix(contentOut, 'w');
end

function contentOut = annotateSplitUnresolvedForPrefix(contentIn, prefix)
contentOut = contentIn;
tOpenPattern = ['<' prefix ':t(?:\s+[^>]*)?>'];
tClosePattern = ['</' prefix ':t>'];
tCloseTag = ['</' prefix ':t>'];
rPrPattern = ['(?:<' prefix ':rPr[^>]*/>\s*|<' prefix ':rPr[^>]*>(?:\s*<[^>]+>\s*)*</' prefix ':rPr>\s*)?'];

splitPattern = ['(' tOpenPattern ')' '\*' tClosePattern ...
    '\s*</' prefix ':r>\s*<' prefix ':r[^>]*>\s*' ...
    rPrPattern ...
    tOpenPattern '((?![Ff]igure\b)[A-Za-z][A-Za-z0-9_]*)' tClosePattern];
splitReplacement = ['$1*$2 [NA]' tCloseTag];
contentOut = regexprep(contentOut, splitPattern, splitReplacement);
end

function writeTextFile(pathStr, content)
fid = fopen(pathStr, 'w');
if fid < 0
    error('Unable to write file: %s', pathStr);
end
cleanObj = onCleanup(@()fclose(fid));
fwrite(fid, content, 'char');
end

function cleanupTempDir(pathStr)
if isfolder(pathStr)
    try
        rmdir(pathStr, 's');
    catch
    end
end
end

function ok = replaceTextPlaceholdersInWordViaCom(docPath, placeholderMap)
ok = false;
if ~ispc
    return;
end
wordApp = [];
docObj = [];
try
    wordApp = actxserver('Word.Application');
    docObj = wordApp.Documents.Open(docPath, false, false, false);
    keys = placeholderMap.keys;
    for i = 1:numel(keys)
        token = string(keys{i});
        replacement = string(placeholderMap(keys{i}));
        replaceWordTokenEverywhere(docObj, token, replacement);
    end
    replaceWordShapeTextPlaceholders(docObj, placeholderMap);
    docObj.Save;
    ok = true;
catch
    ok = false;
end
cleanupWordAutomation(wordApp, docObj);
end

function replaceFigurePlaceholdersInWord(docPath, plotConfig, context, fileLabel, colorMap)
if ~ispc || isempty(plotConfig)
    return;
end

wordApp = [];
docObj = [];
imageMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
try
    wordApp = actxserver('Word.Application');
    docObj = wordApp.Documents.Open(docPath, false, false, false);
    try
        allText = string(docObj.Content.Text);
        tokens = getFigurePlaceholderTokensFromText(allText);
        if ~isempty(tokens)
            for i = 1:numel(tokens)
                try
                    token = string(tokens(i));
                    [~, figNo, subNo, found] = parseFigurePlaceholder(token, plotConfig);
                    if ~found
                        continue;
                    end
                    cacheKey = buildFigureCacheKey(figNo, subNo);
                    if isKey(imageMap, cacheKey)
                        imgPath = string(imageMap(cacheKey));
                    else
                        imgPath = exportFigurePlaceholderImage(plotConfig, context, fileLabel, colorMap, figNo, subNo);
                        if strlength(imgPath) == 0
                            continue;
                        end
                        cacheImagePath(imageMap, cacheKey, imgPath);
                    end
                    replaceWordTokenWithImage(docObj, token, imgPath);
                catch
                end
            end
        end
    catch
    end

    % Also process figure placeholders inside Word Text Boxes (Shape text).
    try
        replaceWordShapeFigurePlaceholders(docObj, plotConfig, context, fileLabel, colorMap, imageMap);
    catch
    end
    docObj.Save;
catch ME
    warning('Word figure placeholder replacement failed for %s: %s', docPath, ME.message);
end
cleanupWordFigureAutomation(wordApp, docObj, imageMap);
end

function replaceWordTokenEverywhere(docObj, tokenText, replacementText)
docEnd = docObj.Content.End;
searchRange = docObj.Range(0, docEnd);
while true
    findObj = searchRange.Find;
    configureWordFind(findObj, tokenText);
    found = logical(findObj.Execute);
    if ~found
        break;
    end
    searchRange.Text = char(replacementText);
    nextStart = searchRange.End;
    docEnd = docObj.Content.End;
    if nextStart >= docEnd
        break;
    end
    searchRange = docObj.Range(nextStart, docEnd);
end
end

function replaceWordTokenWithImage(docObj, tokenText, imgPath)
docEnd = docObj.Content.End;
searchRange = docObj.Range(0, docEnd);
while true
    findObj = searchRange.Find;
    configureWordFind(findObj, tokenText);
    found = logical(findObj.Execute);
    if ~found
        break;
    end

    try
        insertRange = searchRange.Duplicate;
        searchRange.Text = '';
        insertRange.InlineShapes.AddPicture(char(imgPath));
    catch
    end

    nextStart = searchRange.End;
    docEnd = docObj.Content.End;
    if nextStart >= docEnd
        break;
    end
    searchRange = docObj.Range(nextStart, docEnd);
end
end

function replaceWordShapeFigurePlaceholders(docObj, plotConfig, context, fileLabel, colorMap, imageMap)
shapeCollections = getWordShapeCollections(docObj);
for iCollection = 1:numel(shapeCollections)
    shpCollection = shapeCollections{iCollection};
    try
        shapeCount = shpCollection.Count;
    catch
        shapeCount = 0;
    end
    for iShape = shapeCount:-1:1
        try
            shp = shpCollection.Item(iShape);
            processWordShapeFigurePlaceholder(shp, plotConfig, context, fileLabel, colorMap, imageMap);
        catch
        end
    end
end
end

function processWordShapeFigurePlaceholder(shp, plotConfig, context, fileLabel, colorMap, imageMap)
% Recurse grouped shapes.
try
    groupCount = shp.GroupItems.Count;
catch
    groupCount = 0;
end
if groupCount > 0
    for iGroup = groupCount:-1:1
        try
            processWordShapeFigurePlaceholder(shp.GroupItems.Item(iGroup), plotConfig, context, fileLabel, colorMap, imageMap);
        catch
        end
    end
    return;
end

rawText = getWordShapeText(shp);
if strlength(rawText) == 0
    return;
end

tokens = getFigurePlaceholderTokensFromText(rawText);
if isempty(tokens)
    return;
end

% Use first valid figure token found in the shape text.
figNo = NaN;
subNo = NaN;
found = false;
for iTok = 1:numel(tokens)
    [~, figNoTry, subNoTry, ok] = parseFigurePlaceholder(tokens(iTok), plotConfig);
    if ok
        figNo = figNoTry;
        subNo = subNoTry;
        found = true;
        break;
    end
end
if ~found
    markWordShapeUnresolved(shp, rawText);
    return;
end

cacheKey = buildFigureCacheKey(figNo, subNo);
if isKey(imageMap, cacheKey)
    imgPath = string(imageMap(cacheKey));
else
    imgPath = exportFigurePlaceholderImage(plotConfig, context, fileLabel, colorMap, figNo, subNo);
    if strlength(imgPath) == 0
        markWordShapeUnresolved(shp, rawText);
        return;
    end
    cacheImagePath(imageMap, cacheKey, imgPath);
end

% Keep the original Text Box/shape in place and fill it with image.
% This preserves page layout and avoids moving content.
try
    placed = tryFillWordShapeWithImage(shp, imgPath);
catch
    placed = false;
end
if ~placed
    markWordShapeUnresolved(shp, rawText);
end
end

function replaceWordShapeTextPlaceholders(docObj, placeholderMap)
shapeCollections = getWordShapeCollections(docObj);
for iCollection = 1:numel(shapeCollections)
    shpCollection = shapeCollections{iCollection};
    try
        shapeCount = shpCollection.Count;
    catch
        shapeCount = 0;
    end
    for iShape = shapeCount:-1:1
        try
            shp = shpCollection.Item(iShape);
            processWordShapeTextPlaceholder(shp, placeholderMap);
        catch
        end
    end
end
end

function processWordShapeTextPlaceholder(shp, placeholderMap)
% Recurse grouped shapes.
try
    groupCount = shp.GroupItems.Count;
catch
    groupCount = 0;
end
if groupCount > 0
    for iGroup = groupCount:-1:1
        try
            processWordShapeTextPlaceholder(shp.GroupItems.Item(iGroup), placeholderMap);
        catch
        end
    end
    return;
end

rawText = getWordShapeText(shp);
if strlength(rawText) == 0
    return;
end

updatedText = rawText;
keys = placeholderMap.keys;
for i = 1:numel(keys)
    token = string(keys{i});
    value = string(placeholderMap(keys{i}));
    updatedText = strrep(updatedText, token, value);
end

if strcmp(char(updatedText), char(rawText))
    return;
end
setWordShapeText(shp, updatedText);
end

function shapeCollections = getWordShapeCollections(docObj)
shapeCollections = {};
try
    shapeCollections{end + 1} = docObj.Shapes;
catch
end

try
    sectionCount = docObj.Sections.Count;
catch
    sectionCount = 0;
end
for iSection = 1:sectionCount
    try
        secObj = docObj.Sections.Item(iSection);
    catch
        continue;
    end

    try
        headerCount = secObj.Headers.Count;
    catch
        headerCount = 0;
    end
    for iHeader = 1:headerCount
        try
            shapeCollections{end + 1} = secObj.Headers.Item(iHeader).Shapes; %#ok<AGROW>
        catch
        end
    end

    try
        footerCount = secObj.Footers.Count;
    catch
        footerCount = 0;
    end
    for iFooter = 1:footerCount
        try
            shapeCollections{end + 1} = secObj.Footers.Item(iFooter).Shapes; %#ok<AGROW>
        catch
        end
    end
end
end

function txt = getWordShapeText(shp)
txt = "";
% Modern Word text boxes expose content through TextFrame2.
try
    hasTextFrame2 = logical(shp.TextFrame2.HasText);
catch
    hasTextFrame2 = false;
end
if hasTextFrame2
    try
        txt2 = string(shp.TextFrame2.TextRange.Text);
        if strlength(strtrim(erase(txt2, char(13)))) > 0
            txt = txt2;
            return;
        end
    catch
    end
end

% Fallback for legacy shapes.
try
    hasTextFrame = logical(shp.TextFrame.HasText);
catch
    hasTextFrame = false;
end
if ~hasTextFrame
    return;
end

try
    txt = string(shp.TextFrame.TextRange.Text);
    % Some textbox variants expose "oTextBox" here; treat as non-content.
    if strcmpi(strtrim(char(txt)), 'oTextBox')
        txt = "";
    end
catch
    txt = "";
end
end

function setWordShapeText(shp, txt)
value = char(string(txt));
try
    shp.TextFrame.TextRange.Text = value;
    return;
catch
end
try
    shp.TextFrame2.TextRange.Text = value;
catch
end
end

function ok = tryFillWordShapeWithImage(shp, imgPath)
ok = false;
if ~isfile(imgPath)
    return;
end
try
    shp.Fill.UserPicture(char(imgPath));
catch
    return;
end
setWordShapeText(shp, "");
ok = true;
end

function markWordShapeUnresolved(shp, rawText)
if contains(rawText, "[NA]")
    return;
end
setWordShapeText(shp, rawText + " [NA]");
end

function configureWordFind(findObj, tokenText)
try
    findObj.ClearFormatting;
catch
end
try
    findObj.Replacement.ClearFormatting;
catch
end
findObj.Text = char(string(tokenText));
findObj.Forward = true;
findObj.Wrap = 0; % wdFindStop
findObj.Format = false;
findObj.MatchCase = false;
findObj.MatchWholeWord = false;
findObj.MatchWildcards = false;
findObj.MatchSoundsLike = false;
findObj.MatchAllWordForms = false;
end

function tokens = getFigurePlaceholderTokensFromText(rawText)
txt = normalizeWordPlaceholderText(rawText);
matches = regexp(txt, '(?i)\*\*\s*Figure\s*\[\s*[^\]\r\n]+\s*\](?:\s*\[\s*\d+\s*\])?', 'match');
if isempty(matches)
    tokens = strings(0, 1);
    return;
end
tokens = unique(string(matches), 'stable');
tokens = tokens(:);
end

function out = normalizeWordPlaceholderText(rawText)
txt = char(string(rawText));
txt = strrep(txt, char(160), ' ');
txt = strrep(txt, char(8203), '');
txt = strrep(txt, char(8204), '');
txt = strrep(txt, char(8205), '');
txt = regexprep(txt, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '');
out = string(txt);
end

function cleanupWordAutomation(wordApp, docObj)
try
    if ~isempty(docObj)
        try
            docObj.Close(false);
        catch
        end
    end
catch
end
try
    if ~isempty(wordApp)
        try
            wordApp.Quit;
        catch
        end
        try
            delete(wordApp);
        catch
        end
    end
catch
end
end

function cleanupWordFigureAutomation(wordApp, docObj, imageMap)
try
    keys = imageMap.keys;
    for i = 1:numel(keys)
        p = imageMap(keys{i});
        if isfile(p)
            delete(p);
        end
    end
catch
end
cleanupWordAutomation(wordApp, docObj);
end

function replaceFigurePlaceholdersInPpt(pptPath, plotConfig, context, fileLabel, colorMap)
if ~ispc
    return;
end
if isempty(plotConfig)
    return;
end

% Use PowerShell automation first (more robust across MATLAB COM environments).
okPs = replaceFigurePlaceholdersInPptViaPowerShell(pptPath, plotConfig, context, fileLabel, colorMap);
if okPs
    return;
end

try
    pptApp = actxserver('PowerPoint.Application');
    % Some environments reject hiding PowerPoint window; avoid forcing visibility.
catch ME
    warning('MATLAB:FigurePlaceholderReplaceUnavailable', '%s', ...
        ['Figure placeholders were not replaced. PowerShell path failed and MATLAB COM is unavailable (' ME.message ').']);
    return;
end

imageMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
cleanupObj = onCleanup(@()cleanupPptAutomation(pptApp, imageMap));

presentation = [];
try
    presentation = pptApp.Presentations.Open(pptPath, false, false, false);
    slideCount = presentation.Slides.Count;
    for iSlide = 1:slideCount
        slideObj = presentation.Slides.Item(iSlide);
        try
            replaceFigurePlaceholdersInSlide(slideObj, plotConfig, context, fileLabel, colorMap, imageMap);
        catch slideME
            warning('Figure placeholder replacement failed on slide %d: %s', iSlide, slideME.message);
        end
    end
    presentation.Save;
    presentation.Close;
catch ME
    if ~isempty(presentation)
        try
            presentation.Close;
        catch
        end
    end
    warning('Failed while replacing figure placeholders in %s: %s (%s)', pptPath, ME.message, ME.identifier);
end
end

function replaceFigurePlaceholdersInSlide(slideObj, plotConfig, context, fileLabel, colorMap, imageMap)
shapeCount = slideObj.Shapes.Count;
for iShape = shapeCount:-1:1
    shp = slideObj.Shapes.Item(iShape);

    % Handle table cells
    hasTable = false;
    try
        hasTable = logical(shp.HasTable);
    catch
    end
    if hasTable
        tbl = shp.Table;
        for r = 1:tbl.Rows.Count
            for c = 1:tbl.Columns.Count
                cellShape = tbl.Cell(r, c).Shape;
                processFigurePlaceholderShape(slideObj, cellShape, false, ...
                    plotConfig, context, fileLabel, colorMap, imageMap);
            end
        end
    end

    % Handle regular text shapes
    processFigurePlaceholderShape(slideObj, shp, true, ...
        plotConfig, context, fileLabel, colorMap, imageMap);
end
end

function processFigurePlaceholderShape(slideObj, shp, canDelete, plotConfig, context, fileLabel, colorMap, imageMap)
rawText = getShapeText(shp);
if strlength(rawText) == 0
    return;
end

    [token, figNo, subNo, found] = parseFigurePlaceholder(rawText, plotConfig);
if ~found
    return;
end

cacheKey = buildFigureCacheKey(figNo, subNo);
if isKey(imageMap, cacheKey)
    imgPath = imageMap(cacheKey);
else
    imgPath = exportFigurePlaceholderImage(plotConfig, context, fileLabel, colorMap, figNo, subNo);
    if strlength(imgPath) == 0
        markShapeUnresolved(shp, rawText);
        return;
    end
    cacheImagePath(imageMap, cacheKey, imgPath);
end

left = shp.Left;
top = shp.Top;
width = shp.Width;
height = shp.Height;

added = false;
try
    % LinkToFile must be msoFalse(0), SaveWithDocument must be msoTrue(-1).
    slideObj.Shapes.AddPicture(char(imgPath), 0, -1, double(left), double(top), double(width), double(height));
    added = true;
catch
end

if ~added
    try
        % Fallback: insert then size to placeholder bounds.
        newShape = slideObj.Shapes.AddPicture(char(imgPath), 0, -1, double(left), double(top));
        newShape.LockAspectRatio = 0;
        newShape.Width = double(width);
        newShape.Height = double(height);
        added = true;
    catch ME
        markShapeUnresolved(shp, rawText);
        warning('Could not place figure image for placeholder "%s": %s', token, ME.message);
        return;
    end
end

if canDelete && added
    try
        shp.Delete;
        return;
    catch
    end
end

setShapeText(shp, '');
end

function [token, figNo, subNo, found] = parseFigurePlaceholder(rawText, plotConfig)
token = "";
figNo = NaN;
subNo = NaN;
found = false;

txt = char(normalizeWordPlaceholderText(rawText));
tokenMatch = regexp(txt, '(?i)\*\*\s*Figure\s*\[\s*[^\]]+\s*\](?:\s*\[\s*\d+\s*\])?', 'match', 'once');
if isempty(tokenMatch)
    return;
end
token = string(tokenMatch);

figRefParts = regexp(tokenMatch, '(?i)^\s*\*\*\s*Figure\s*\[\s*([^\]]+)\s*\]', 'tokens', 'once');
if isempty(figRefParts)
    return;
end

figRef = string(figRefParts{1});
[figNo, okFig] = resolveFigureReference(figRef, plotConfig);
if ~okFig
    return;
end

subParts = regexp(tokenMatch, '\]\s*\[\s*(\d+)\s*\]\s*$', 'tokens', 'once');
if ~isempty(subParts)
    subNo = str2double(string(subParts{1}));
    if isnan(subNo)
        return;
    end
end

found = true;
end

function [figNo, ok] = resolveFigureReference(figRef, plotConfig)
figNo = NaN;
ok = false;

if ~ismember("Figure No", string(plotConfig.Properties.VariableNames))
    return;
end

figRef = strtrim(string(figRef));
if strlength(figRef) == 0 || ismissing(figRef)
    return;
end

numCandidate = str2double(figRef);
figureNos = unique(double(plotConfig.("Figure No")), 'stable');
if ~isnan(numCandidate) && any(figureNos == numCandidate)
    figNo = numCandidate;
    ok = true;
    return;
end

nameCols = getFigureNameColumns(plotConfig);
if isempty(nameCols)
    return;
end

refNorm = normalizeFigureReferenceText(figRef);
if strlength(refNorm) == 0
    return;
end

for iCol = 1:numel(nameCols)
    colName = nameCols(iCol);
    colVals = strtrim(string(plotConfig.(colName)));
    valid = ~ismissing(colVals) & strlength(colVals) > 0;
    if ~any(valid)
        continue;
    end
    normVals = normalizeFigureReferenceText(colVals(valid));
    hit = find(normVals == refNorm, 1, 'first');
    if isempty(hit)
        continue;
    end
    validIdx = find(valid);
    figNo = double(plotConfig.("Figure No")(validIdx(hit)));
    ok = true;
    return;
end
end

function nameCols = getFigureNameColumns(plotConfig)
known = ["Figure Name", "Figure", "Figure Title", "Name"];
vars = string(plotConfig.Properties.VariableNames);
nameCols = known(ismember(known, vars));
nameCols = nameCols(:);
end

function out = normalizeFigureReferenceText(inText)
out = lower(strtrim(string(inText)));
out = regexprep(out, '\s+', ' ');
out = regexprep(out, '[^a-z0-9_ ]', '');
out = regexprep(out, '\s+', '_');
out = regexprep(out, '^_+|_+$', '');
end

function cacheKey = buildFigureCacheKey(figNo, subNo)
if isnan(subNo)
    cacheKey = sprintf('F%d', figNo);
else
    cacheKey = sprintf('F%d_S%d', figNo, subNo);
end
end

function cacheImagePath(imageMap, cacheKey, imgPath)
imageMap(char(cacheKey)) = char(imgPath); %#ok<NASGU>
end

function imagePath = exportFigurePlaceholderImage(plotConfig, context, fileLabel, colorMap, figNo, subNo)
imagePath = "";
figRows = getFigureRowsForPlaceholder(plotConfig, figNo);
if isempty(figRows)
    return;
end

subplotNos = unique(figRows.("Subplot"), 'stable');
if isnan(subNo)
    targetSubplots = subplotNos;
else
    targetSubplots = subplotNos(subplotNos == subNo);
end
if isempty(targetSubplots)
    return;
end

figTitle = sprintf('%s - Figure %d', string(fileLabel), figNo);
hf = figure('Visible', 'off', 'Color', 'w', 'Position', [100 100 1280 720]);
cleanupFig = onCleanup(@()closeIfValid(hf));

t = tiledlayout(numel(targetSubplots), 1, 'TileSpacing', 'compact', 'Padding', 'compact');
title(t, figTitle, 'Interpreter', 'none');

for iSub = 1:numel(targetSubplots)
    ax = nexttile(t, iSub);
    hold(ax, 'on');
    grid(ax, 'on');

    subRows = figRows(figRows.("Subplot") == targetSubplots(iSub), :);
    axisKinds = upper(strtrim(string(subRows.("Axis"))));

    titleRows = subRows(axisKinds == "T", :);
    if ~isempty(titleRows)
        subTitle = strtrim(string(titleRows.("Signals and Titles")(1)));
        if strlength(subTitle) > 0 && ~ismissing(subTitle)
            title(ax, subTitle, 'Interpreter', 'none');
        end
    end

    xRows = subRows(axisKinds == "X", :);
    xData = [];
    hasXData = false;
    if ~isempty(xRows)
        xExpr = strtrim(string(xRows.("Signals and Titles")(1)));
        if strlength(xExpr) > 0 && ~ismissing(xExpr)
            [xCandidate, okX] = tryEvaluateEquation(xExpr, context);
            if okX && isnumeric(xCandidate)
                xData = xCandidate;
                hasXData = true;
            end
        end
        xLabelText = buildAxisLabel(xRows.("Label")(1), xRows.("Unit")(1));
        if strlength(xLabelText) > 0
            xlabel(ax, xLabelText, 'Interpreter', 'none');
        end
    end

    yRows = subRows(axisKinds == "Y", :);
    anyLinePlotted = false;
    hasLeftYLabel = false;
    hasRightYLabel = false;

    for iY = 1:height(yRows)
        if ~shouldPlotRow(yRows(iY, :), false)
            continue;
        end

        yExpr = strtrim(string(yRows.("Signals and Titles")(iY)));
        if strlength(yExpr) == 0 || ismissing(yExpr)
            continue;
        end

        [yData, okY] = tryEvaluateEquation(yExpr, context);
        if ~okY || ~isnumeric(yData) || isempty(yData)
            continue;
        end

        axisPos = upper(strtrim(string(yRows.("Axis Pos")(iY))));
        if ismissing(axisPos) || strlength(axisPos) == 0
            axisPos = "L";
        end
        if axisPos == "R"
            yyaxis(ax, 'right');
            side = "R";
        else
            yyaxis(ax, 'left');
            side = "L";
        end

        styleText = strtrim(string(yRows.("Style")(iY)));
        hasStyle = strlength(styleText) > 0 && ~ismissing(styleText);
        if hasStyle
            styleArg = char(styleText);
        end

        lineColor = resolvePlotColor(yRows.("Color")(iY), colorMap);
        lineWidth = resolveLineWidth(yRows.("Width")(iY));

        if hasXData && numel(xData) == numel(yData)
            if hasStyle
                p = plot(ax, xData, yData, styleArg, 'LineWidth', lineWidth, 'Color', lineColor);
            else
                p = plot(ax, xData, yData, 'LineWidth', lineWidth, 'Color', lineColor);
            end
        else
            if hasStyle
                p = plot(ax, yData, styleArg, 'LineWidth', lineWidth, 'Color', lineColor);
            else
                p = plot(ax, yData, 'LineWidth', lineWidth, 'Color', lineColor);
            end
        end
        anyLinePlotted = true;

        legendText = "";
        if ismember("Legend", string(yRows.Properties.VariableNames))
            legendText = strtrim(string(yRows.("Legend")(iY)));
        end
        if strlength(legendText) > 0 && ~ismissing(legendText)
            p.DisplayName = legendText;
        else
            p.DisplayName = yExpr;
        end

        yLabelText = buildAxisLabel(yRows.("Label")(iY), yRows.("Unit")(iY));
        if strlength(yLabelText) > 0
            if side == "L" && ~hasLeftYLabel
                ylabel(ax, yLabelText, 'Interpreter', 'none');
                hasLeftYLabel = true;
            elseif side == "R" && ~hasRightYLabel
                ylabel(ax, yLabelText, 'Interpreter', 'none');
                hasRightYLabel = true;
            end
        end
    end

    if anyLinePlotted
        legend(ax, 'show', 'Location', 'best', 'Interpreter', 'none');
    end
    hold(ax, 'off');
end

imagePath = string([tempname '.png']);
exportgraphics(hf, imagePath, 'Resolution', 200);
end

function figRows = getFigureRowsForPlaceholder(plotConfig, figNo)
hasGroupCol = ismember("Group", string(plotConfig.Properties.VariableNames));
if hasGroupCol
    groupsAll = normalizeGroupValues(plotConfig.("Group"));
    rows = plotConfig(plotConfig.("Figure No") == figNo, :);
    if isempty(rows)
        figRows = rows;
        return;
    end
    rowsGroups = groupsAll(plotConfig.("Figure No") == figNo);
    firstGroup = rowsGroups(1);
    figRows = rows(rowsGroups == firstGroup, :);
else
    figRows = plotConfig(plotConfig.("Figure No") == figNo, :);
end
end

function txt = getShapeText(shp)
txt = "";
try
    hasTextFrame = logical(shp.HasTextFrame);
catch
    hasTextFrame = false;
end
if ~hasTextFrame
    return;
end
try
    hasText = logical(shp.TextFrame.HasText);
catch
    hasText = false;
end
if ~hasText
    return;
end
try
    txt = string(shp.TextFrame.TextRange.Text);
catch
    txt = "";
end
end

function setShapeText(shp, txt)
try
    shp.TextFrame.TextRange.Text = char(string(txt));
catch
end
end

function markShapeUnresolved(shp, rawText)
if contains(rawText, '[NA]')
    return;
end
setShapeText(shp, rawText + " [NA]");
end

function closeIfValid(h)
try
    if ishghandle(h)
        close(h);
    end
catch
end
end

function cleanupPptAutomation(pptApp, imageMap)
% Remove temporary images.
try
    keys = imageMap.keys;
    for i = 1:numel(keys)
        p = imageMap(keys{i});
        if isfile(p)
            delete(p);
        end
    end
catch
end

% Close PowerPoint automation instance.
try
    pptApp.Quit;
catch
end
try
    delete(pptApp);
catch
end
end

function ok = replaceFigurePlaceholdersInPptViaPowerShell(pptPath, plotConfig, context, fileLabel, colorMap)
ok = false;
if ~ispc
    return;
end

figureNos = unique(plotConfig.("Figure No"), 'stable');
if isempty(figureNos)
    ok = true;
    return;
end

defaultFig = double(figureNos(1));
imgMap = containers.Map('KeyType', 'char', 'ValueType', 'char');
tempAssets = cell(0, 1);
nTempAssets = 0;
try
    requestedKeys = getRequestedFigureKeysFromPpt(pptPath, plotConfig);
    if isempty(requestedKeys)
        requestedKeys = getAllFigureKeys(plotConfig);
    end

    tempAssets = cell(numel(requestedKeys) + 4, 1);

    for iKey = 1:numel(requestedKeys)
        [figNo, subNo, parsed] = parseFigureCacheKey(requestedKeys(iKey));
        if ~parsed
            continue;
        end
        keyText = char(requestedKeys(iKey));
        if isKey(imgMap, keyText)
            continue;
        end
        imgPath = exportFigurePlaceholderImage(plotConfig, context, fileLabel, colorMap, figNo, subNo);
        if strlength(imgPath) > 0
            imgMap(keyText) = char(imgPath);
            nTempAssets = nTempAssets + 1;
            tempAssets{nTempAssets} = char(imgPath);
        end
    end

    if isempty(imgMap.keys)
        cleanupFiles(tempAssets(1:nTempAssets));
        return;
    end

    mapStruct = struct();
    mapKeys = imgMap.keys;
    for i = 1:numel(mapKeys)
        mapStruct.(mapKeys{i}) = imgMap(mapKeys{i});
        [figNoAlias, subNoAlias, okAlias] = parseFigureCacheKey(mapKeys{i});
        if okAlias
            aliasKeys = getFigureNameCacheKeys(plotConfig, figNoAlias, subNoAlias);
            for iAlias = 1:numel(aliasKeys)
                aliasKey = char(aliasKeys(iAlias));
                mapStruct.(aliasKey) = imgMap(mapKeys{i});
            end
        end
    end

    mapJsonPath = [tempname '.json'];
    psPath = [tempname '.ps1'];
    nTempAssets = nTempAssets + 1;
    tempAssets{nTempAssets} = mapJsonPath;
    nTempAssets = nTempAssets + 1;
    tempAssets{nTempAssets} = psPath;

    writeTextFile(mapJsonPath, jsonencode(mapStruct));
    writeTextFile(psPath, buildPptFallbackPowerShellScript());

    cmd = sprintf('powershell -NoProfile -ExecutionPolicy Bypass -File "%s" -pptPath "%s" -mapJson "%s" -defaultFig %d', ...
        psPath, pptPath, mapJsonPath, defaultFig);
    [status, out] = system(cmd);
    if status == 0
        ok = true;
    else
        warning('PowerShell fallback failed: %s', strtrim(out));
    end
catch ME
    warning('MATLAB:PowerShellFallbackException', '%s', ...
        ['PowerShell fallback threw an exception: ' ME.message]);
end
cleanupFiles(tempAssets(1:nTempAssets));
end

function tf = hasFigurePlaceholdersInPpt(pptPath)
tf = false;
[~, ~, ext] = fileparts(pptPath);
ext = lower(string(ext));
if ext ~= ".pptx" && ext ~= ".pptm"
    return;
end

zipPath = [tempname '.zip'];
workDir = tempname;
try
    copyfile(pptPath, zipPath, 'f');
    unzip(zipPath, workDir);
    xmlFiles = dir(fullfile(workDir, 'ppt', 'slides', '*.xml'));
    for i = 1:numel(xmlFiles)
        xmlPath = fullfile(xmlFiles(i).folder, xmlFiles(i).name);
        txt = fileread(xmlPath);
        if ~isempty(regexp(txt, '\*\*Figure', 'once'))
            tf = true;
            break;
        end
    end
catch
    tf = true;
end
cleanupFiles({zipPath});
cleanupTempDir(workDir);
end

function requestedKeys = getRequestedFigureKeysFromPpt(pptPath, plotConfig)
requestedKeys = strings(0, 1);
zipPath = [tempname '.zip'];
workDir = tempname;
figureNos = unique(double(plotConfig.("Figure No")), 'stable');

try
    copyfile(pptPath, zipPath, 'f');
    unzip(zipPath, workDir);
    xmlFiles = dir(fullfile(workDir, 'ppt', 'slides', '*.xml'));
    chunks = cell(numel(xmlFiles), 1);
    nChunks = 0;
    for i = 1:numel(xmlFiles)
        xmlPath = fullfile(xmlFiles(i).folder, xmlFiles(i).name);
        txt = fileread(xmlPath);
        starts = regexp(txt, '(?i)\*\*\s*Figure', 'start');
        if isempty(starts)
            continue;
        end
        localKeys = strings(numel(starts), 1);
        nLocal = 0;
        for k = 1:numel(starts)
            subNo = NaN;
            tail = txt(starts(k):min(numel(txt), starts(k) + 120));
            tokenText = regexprep(tail, '<[^>]+>', '');
            tokenText = regexp(tokenText, '^[^\r\n]*', 'match', 'once');
            tokenMatch = regexp(tokenText, '(?i)\*\*\s*Figure\s*\[\s*[^\]]+\s*\](?:\s*\[\s*\d+\s*\])?', 'match', 'once');
            if isempty(tokenMatch)
                continue;
            end
            figRefParts = regexp(tokenMatch, '(?i)^\s*\*\*\s*Figure\s*\[\s*([^\]]+)\s*\]', 'tokens', 'once');
            if isempty(figRefParts)
                continue;
            end
            [figNo, okFig] = resolveFigureReference(string(figRefParts{1}), plotConfig);
            if ~okFig
                continue;
            end
            subParts = regexp(tokenMatch, '\]\s*\[\s*(\d+)\s*\]\s*$', 'tokens', 'once');
            if ~isempty(subParts)
                subNo = str2double(string(subParts{1}));
                if isnan(subNo)
                    continue;
                end
            end
            if ~any(figureNos == figNo)
                continue;
            end
            nLocal = nLocal + 1;
            localKeys(nLocal, 1) = string(buildFigureCacheKey(figNo, subNo));
        end
        if nLocal > 0
            nChunks = nChunks + 1;
            chunks{nChunks} = localKeys(1:nLocal);
        end
    end
    if nChunks > 0
        requestedKeys = vertcat(chunks{1:nChunks});
    end
    requestedKeys = unique(requestedKeys, 'stable');
catch
    requestedKeys = strings(0, 1);
end

cleanupFiles({zipPath});
cleanupTempDir(workDir);
end

function keys = getAllFigureKeys(plotConfig)
figureNos = unique(plotConfig.("Figure No"), 'stable');
subNosByFig = cell(numel(figureNos), 1);
totalKeys = 0;
for iFig = 1:numel(figureNos)
    figNo = double(figureNos(iFig));
    figRows = getFigureRowsForPlaceholder(plotConfig, figNo);
    subNos = unique(figRows.("Subplot"), 'stable');
    subNosByFig{iFig} = subNos;
    totalKeys = totalKeys + 1 + numel(subNos);
end

keys = strings(totalKeys, 1);
idx = 0;
for iFig = 1:numel(figureNos)
    figNo = double(figureNos(iFig));
    idx = idx + 1;
    keys(idx, 1) = string(buildFigureCacheKey(figNo, NaN));
    subNos = subNosByFig{iFig};
    for iSub = 1:numel(subNos)
        idx = idx + 1;
        keys(idx, 1) = string(buildFigureCacheKey(figNo, double(subNos(iSub))));
    end
end
keys = unique(keys, 'stable');
end

function [figNo, subNo, ok] = parseFigureCacheKey(keyText)
figNo = NaN;
subNo = NaN;
ok = false;
tok = regexp(char(string(keyText)), '^F(\d+)(?:_S(\d+))?$', 'tokens', 'once');
if isempty(tok)
    return;
end
figNo = str2double(tok{1});
if numel(tok) >= 2 && ~isempty(tok{2})
    subNo = str2double(tok{2});
end
ok = ~isnan(figNo);
end

function aliasKeys = getFigureNameCacheKeys(plotConfig, figNo, subNo)
aliasKeys = strings(0, 1);
if ~ismember("Figure No", string(plotConfig.Properties.VariableNames))
    return;
end

nameCols = getFigureNameColumns(plotConfig);
if isempty(nameCols)
    return;
end

rows = plotConfig(double(plotConfig.("Figure No")) == double(figNo), :);
if isempty(rows)
    return;
end

keys = strings(0, 1);
for iCol = 1:numel(nameCols)
    colName = nameCols(iCol);
    names = strtrim(string(rows.(colName)));
    names = names(~ismissing(names) & strlength(names) > 0);
    for iName = 1:numel(names)
        key = buildFigureNameCacheKey(names(iName), subNo);
        if strlength(key) > 0
            keys(end+1, 1) = key; %#ok<AGROW>
        end
    end
end
if ~isempty(keys)
    aliasKeys = unique(keys, 'stable');
end
end

function key = buildFigureNameCacheKey(figName, subNo)
normName = normalizeFigureReferenceText(figName);
if strlength(normName) == 0
    key = "";
    return;
end
if isnan(subNo)
    key = "N_" + normName;
else
    key = string(sprintf('N_%s_S%d', char(normName), double(subNo)));
end
end

function scriptText = buildPptFallbackPowerShellScript()
lines = {
'param('
'  [Parameter(Mandatory=$true)][string]$pptPath,'
'  [Parameter(Mandatory=$true)][string]$mapJson,'
'  [Parameter(Mandatory=$true)][int]$defaultFig'
')'
''
'$ErrorActionPreference = ''Stop'''
'$mapObj = Get-Content -Raw -Path $mapJson | ConvertFrom-Json'
''
'function Get-MapValue($obj, [string]$key) {'
'  try {'
'    if ($null -eq $obj) { return $null }'
'    $prop = $obj.PSObject.Properties[$key]'
'    if ($null -ne $prop) { return [string]$prop.Value }'
'  } catch {}'
'  return $null'
'}'
''
'function Get-ShapeText($shape) {'
'  try {'
'    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {'
'      return [string]$shape.TextFrame.TextRange.Text'
'    }'
'  } catch {}'
'  return '''''
'}'
''
'function Set-ShapeText($shape, [string]$txt) {'
'  try { $shape.TextFrame.TextRange.Text = $txt } catch {}'
'}'
''
'function Process-Shape($slide, $shape) {'
'  $txt = Get-ShapeText $shape'
'  if ([string]::IsNullOrEmpty($txt)) { return }'
''
'  $m = [regex]::Match($txt, ''(?i)\*\*\s*Figure\s*\[\s*([^\]]+)\s*\](?:\s*\[\s*(\d+)\s*\])?'')'
'  if (-not $m.Success) { return }'
''
'  $figToken = [string]$m.Groups[1].Value'
'  $sub = $null'
'  if ($m.Groups.Count -ge 3 -and $m.Groups[2].Success) {'
'    $subVal = 0'
'    if ([int]::TryParse($m.Groups[2].Value, [ref]$subVal)) { $sub = $subVal } else { return }'
'  }'
''
'  $figVal = 0'
'  if ([int]::TryParse($figToken.Trim(), [ref]$figVal)) {'
'    $keyBase = ''F{0}'' -f $figVal'
'  } else {'
'    $nameNorm = $figToken.ToLowerInvariant().Trim()'
'    $nameNorm = [regex]::Replace($nameNorm, ''\s+'', '' '')'
'    $nameNorm = [regex]::Replace($nameNorm, ''[^a-z0-9_ ]'', '''')'
'    $nameNorm = [regex]::Replace($nameNorm, ''\s+'', ''_'')'
'    $nameNorm = [regex]::Replace($nameNorm, ''^_+|_+$'', '''')'
'    if ([string]::IsNullOrWhiteSpace($nameNorm)) { return }'
'    $keyBase = ''N_{0}'' -f $nameNorm'
'  }'
''
'  $key = if ($null -ne $sub) { ''{0}_S{1}'' -f $keyBase, $sub } else { $keyBase }'
''
'  $imgPath = Get-MapValue $mapObj $key'
'  if (($null -ne $imgPath) -and (Test-Path $imgPath)) {'
'    $left = [double]$shape.Left'
'    $top = [double]$shape.Top'
'    $width = [double]$shape.Width'
'    $height = [double]$shape.Height'
''
'    try {'
'      $null = $slide.Shapes.AddPicture($imgPath, 0, -1, $left, $top, $width, $height)'
'    } catch {'
'      $pic = $slide.Shapes.AddPicture($imgPath, 0, -1, $left, $top)'
'      $pic.LockAspectRatio = 0'
'      $pic.Width = $width'
'      $pic.Height = $height'
'    }'
''
'    try { $shape.Delete(); return } catch {}'
'    Set-ShapeText $shape '''''
'  } else {'
'    if ($txt -notmatch ''\[NA\]'') {'
'      Set-ShapeText $shape ($txt + '' [NA]'')'
'    }'
'  }'
'}'
''
'$ppt = $null'
'$pres = $null'
'try {'
'  $ppt = New-Object -ComObject PowerPoint.Application'
'  try { $ppt.Visible = 1 } catch {}'
'  $pres = $ppt.Presentations.Open($pptPath, $false, $false, $false)'
''
'  foreach ($slide in $pres.Slides) {'
'    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {'
'      $shape = $slide.Shapes.Item($i)'
'      try {'
'        if ($shape.HasTable) {'
'          $tbl = $shape.Table'
'          for ($r = 1; $r -le $tbl.Rows.Count; $r++) {'
'            for ($c = 1; $c -le $tbl.Columns.Count; $c++) {'
'              $cellShape = $tbl.Cell($r, $c).Shape'
'              Process-Shape $slide $cellShape'
'            }'
'          }'
'        }'
'      } catch {}'
''
'      Process-Shape $slide $shape'
'    }'
'  }'
''
'  $pres.Save()'
'  $pres.Close()'
'  $ppt.Quit()'
'  exit 0'
'} catch {'
'  if ($pres -ne $null) { try { $pres.Close() } catch {} }'
'  if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }'
'  Write-Error $_.Exception.Message'
'  exit 1'
'}'
};
scriptText = strjoin(lines, newline);
end

function cleanupFiles(paths)
for i = 1:numel(paths)
    try
        if isfile(paths{i})
            delete(paths{i});
        end
    catch
    end
end
end

function plotForAllFiles(plotConfig, contextsByFile, fileLabels, colorMap, targetGroup, ignorePrintStatus, figSaveDir)
if isempty(plotConfig)
    return;
end

if nargin < 5
    targetGroup = "";
end
if nargin < 6
    ignorePrintStatus = false;
end
if nargin < 7
    figSaveDir = "";
end

figSaveDirsByFile = normalizePerFilePathList(figSaveDir, numel(contextsByFile));
requiredPlotCols = ["Figure No", "Subplot", "Axis", "Signals and Titles"];
for iCol = 1:numel(requiredPlotCols)
    if ~ismember(requiredPlotCols(iCol), string(plotConfig.Properties.VariableNames))
        warning('Plots sheet missing required column "%s". Plotting skipped.', requiredPlotCols(iCol));
        return;
    end
end

hasGroupCol = ismember("Group", string(plotConfig.Properties.VariableNames));
if hasGroupCol
    allGroups = normalizeGroupValues(plotConfig.("Group"));
else
    allGroups = repmat("GENERAL", height(plotConfig), 1);
end

groupNames = unique(allGroups, 'stable');
if strlength(string(targetGroup)) > 0
    targetGroup = string(targetGroup);
    groupNames = groupNames(groupNames == targetGroup);
end

for iFile = 1:numel(contextsByFile)
    context = contextsByFile{iFile};
    fileLabel = string(fileLabels(iFile));
    fileFigSaveDir = "";
    if iFile <= numel(figSaveDirsByFile)
        fileFigSaveDir = string(figSaveDirsByFile{iFile});
    end
    saveFigs = strlength(fileFigSaveDir) > 0 && isfolder(char(fileFigSaveDir));

    for iGroup = 1:numel(groupNames)
        groupMask = allGroups == groupNames(iGroup);
        groupRowsAll = plotConfig(groupMask, :);
        if isempty(groupRowsAll)
            continue;
        end

        figureNos = unique(groupRowsAll.("Figure No"), 'stable');
        for iFig = 1:numel(figureNos)
            figNo = figureNos(iFig);
            figRows = groupRowsAll(groupRowsAll.("Figure No") == figNo, :);
            subplotNos = unique(figRows.("Subplot"), 'stable');
            if isempty(subplotNos) || ~hasPlottableFigureSignals(figRows, ignorePrintStatus)
                continue;
            end

            groupLabel = upper(groupNames(iGroup));
            figTitle = sprintf('Figure %s _ %s _ %s', ...
                char(string(figNo)), char(groupLabel), char(fileLabel));

            hFig = figure('Name', figTitle, 'NumberTitle', 'off', 'Color', 'w', 'Visible', 'off');
            t = tiledlayout(numel(subplotNos), 1, 'TileSpacing', 'compact', 'Padding', 'compact');
            title(t, figTitle, 'Interpreter', 'none');

            for iSub = 1:numel(subplotNos)
                ax = nexttile(t, iSub);
                hold(ax, 'on');
                grid(ax, 'on');

                subRows = figRows(figRows.("Subplot") == subplotNos(iSub), :);
                axisKinds = upper(strtrim(string(subRows.("Axis"))));

                titleRows = subRows(axisKinds == "T", :);
                if ~isempty(titleRows)
                    subTitle = strtrim(string(titleRows.("Signals and Titles")(1)));
                    if strlength(subTitle) > 0 && ~ismissing(subTitle)
                        title(ax, subTitle, 'Interpreter', 'none');
                    end
                end

                xRows = subRows(axisKinds == "X", :);
                xData = [];
                hasXData = false;
                if ~isempty(xRows)
                    xExpr = strtrim(string(xRows.("Signals and Titles")(1)));
                    if strlength(xExpr) > 0 && ~ismissing(xExpr)
                        [xCandidate, okX] = tryEvaluateEquation(xExpr, context);
                        if okX && isnumeric(xCandidate)
                            xData = xCandidate;
                            hasXData = true;
                        end
                    end

                    xLabelText = buildAxisLabel( ...
                        getTableValue(xRows, "Label", 1, ""), ...
                        getTableValue(xRows, "Unit", 1, ""));
                    if strlength(xLabelText) > 0
                        xlabel(ax, xLabelText, 'Interpreter', 'none');
                    end
                end

                yRows = subRows(axisKinds == "Y", :);
                anyLinePlotted = false;
                hasLeftYLabel = false;
                hasRightYLabel = false;

                for iY = 1:height(yRows)
                    if ~shouldPlotRow(yRows(iY, :), ignorePrintStatus)
                        continue;
                    end

                    yExpr = strtrim(string(yRows.("Signals and Titles")(iY)));
                    if strlength(yExpr) == 0 || ismissing(yExpr)
                        continue;
                    end

                    [yData, okY] = tryEvaluateEquation(yExpr, context);
                    if ~okY || ~isnumeric(yData) || isempty(yData)
                        continue;
                    end

                    axisPos = upper(strtrim(string(getTableValue(yRows, "Axis Pos", iY, "L"))));
                    if axisPos == "R"
                        yyaxis(ax, 'right');
                        side = "R";
                    else
                        yyaxis(ax, 'left');
                        side = "L";
                    end

                    styleText = strtrim(string(getTableValue(yRows, "Style", iY, "")));
                    lineColor = resolvePlotColor(getTableValue(yRows, "Color", iY, ""), colorMap);
                    lineWidth = resolveLineWidth(getTableValue(yRows, "Width", iY, 1.2));

                    if strlength(styleText) > 0
                        if hasXData && numel(xData) == numel(yData)
                            p = plot(ax, xData, yData, char(styleText), ...
                                'LineWidth', lineWidth, 'Color', lineColor);
                        else
                            p = plot(ax, yData, char(styleText), ...
                                'LineWidth', lineWidth, 'Color', lineColor);
                        end
                    else
                        if hasXData && numel(xData) == numel(yData)
                            p = plot(ax, xData, yData, 'LineWidth', lineWidth, 'Color', lineColor);
                        else
                            p = plot(ax, yData, 'LineWidth', lineWidth, 'Color', lineColor);
                        end
                    end
                    anyLinePlotted = true;

                    legendText = strtrim(string(getTableValue(yRows, "Legend", iY, "")));
                    if strlength(legendText) > 0 && ~ismissing(legendText)
                        p.DisplayName = legendText;
                    else
                        p.DisplayName = yExpr;
                    end

                    yLabelText = buildAxisLabel( ...
                        getTableValue(yRows, "Label", iY, ""), ...
                        getTableValue(yRows, "Unit", iY, ""));
                    if strlength(yLabelText) > 0
                        if side == "L" && ~hasLeftYLabel
                            ylabel(ax, yLabelText, 'Interpreter', 'none');
                            hasLeftYLabel = true;
                        elseif side == "R" && ~hasRightYLabel
                            ylabel(ax, yLabelText, 'Interpreter', 'none');
                            hasRightYLabel = true;
                        end
                    end
                end

                if anyLinePlotted
                    legend(ax, 'show', 'Location', 'best', 'Interpreter', 'none');
                end
                hold(ax, 'off');
            end

            if saveFigs
                try
                    baseName = sprintf('Figure_%s_%s_%s', ...
                        char(string(figNo)), char(groupNames(iGroup)), char(fileLabel));
                    figPath = makeUniqueFilePath(char(fileFigSaveDir), sanitizeFileName(baseName), '.fig');
                    savefig(hFig, figPath);
                catch saveErr
                    warning('BatchDive:FigureSaveFailed', ...
                        'Failed to save figure %s for %s: %s', ...
                        char(string(figNo)), char(fileLabel), saveErr.message);
                end
            end

            if isgraphics(hFig)
                close(hFig);
            end
        end
    end
end
end

function out = getTableValue(tbl, varName, rowIdx, defaultValue)
if nargin < 4
    defaultValue = "";
end

if ~ismember(varName, string(tbl.Properties.VariableNames))
    out = defaultValue;
    return;
end

out = tbl.(varName)(rowIdx);
if iscell(out) && isscalar(out)
    out = out{1};
end
end

function outPaths = normalizePerFilePathList(pathInput, nFiles)
outPaths = repmat({''}, nFiles, 1);
if nFiles <= 0
    return;
end

if ischar(pathInput) || (isstring(pathInput) && isscalar(pathInput))
    p = char(string(pathInput));
    if strlength(strtrim(string(p))) > 0
        outPaths(:) = {p};
    end
    return;
end

if isstring(pathInput)
    n = min(nFiles, numel(pathInput));
    for i = 1:n
        p = char(string(pathInput(i)));
        if strlength(strtrim(string(p))) > 0
            outPaths{i} = p;
        end
    end
    return;
end

if iscell(pathInput)
    n = min(nFiles, numel(pathInput));
    for i = 1:n
        p = char(string(pathInput{i}));
        if strlength(strtrim(string(p))) > 0
            outPaths{i} = p;
        end
    end
end
end

function out = normalizeGroupValues(groupCol)
out = strtrim(string(groupCol));
out(ismissing(out) | strlength(out) == 0) = "GENERAL";
end

function tf = shouldPlotRow(rowTable, ignorePrintStatus)
tf = false;
vars = string(rowTable.Properties.VariableNames);

if ismember("Axis", vars)
    axisVal = upper(strtrim(string(rowTable.("Axis")(1))));
    if axisVal == "T" || axisVal == "X"
        tf = true;
        return;
    end
end

if nargin >= 2 && ignorePrintStatus
    tf = true;
    return;
end

if ~ismember("Print", vars)
    return;
end

rawVal = rowTable.("Print")(1);
if ismissing(rawVal)
    return;
end

if isstring(rawVal) || ischar(rawVal)
    txt = strtrim(string(rawVal));
    numVal = str2double(txt);
    tf = ~isnan(numVal) && (numVal == 1);
    return;
end

if isnumeric(rawVal) || islogical(rawVal)
    numVal = double(rawVal);
    tf = isfinite(numVal) && (numVal == 1);
    return;
end

txt = strtrim(string(rawVal));
numVal = str2double(txt);
tf = ~isnan(numVal) && (numVal == 1);
end

function tf = hasPlottableFigureSignals(figRows, ignorePrintStatus)
tf = false;
if isempty(figRows) || ~ismember("Axis", string(figRows.Properties.VariableNames))
    return;
end

axisKinds = upper(strtrim(string(figRows.("Axis"))));
yRows = figRows(axisKinds == "Y", :);
if isempty(yRows)
    return;
end

for iRow = 1:height(yRows)
    if shouldPlotRow(yRows(iRow, :), ignorePrintStatus)
        tf = true;
        return;
    end
end
end

function out = buildAxisLabel(labelVal, unitVal)
labelText = strtrim(string(labelVal));
unitText = strtrim(string(unitVal));

if (strlength(labelText) == 0 || ismissing(labelText)) && ...
        (strlength(unitText) == 0 || ismissing(unitText))
    out = "";
elseif strlength(unitText) == 0 || ismissing(unitText)
    out = labelText;
elseif strlength(labelText) == 0 || ismissing(labelText)
    out = unitText;
else
    out = labelText + " (" + unitText + ")";
end
end

function out = resolveLineWidth(widthVal)
out = 1.2;
if isnumeric(widthVal) && isscalar(widthVal) && ~isnan(widthVal)
    out = double(widthVal);
    return;
end

widthText = strtrim(string(widthVal));
if strlength(widthText) > 0 && ~ismissing(widthText)
    w = str2double(widthText);
    if ~isnan(w)
        out = w;
    end
end
end

function out = resolvePlotColor(colorVal, colorMap)
out = [0 0 0];
colorText = lower(strtrim(string(colorVal)));
if strlength(colorText) == 0 || ismissing(colorText)
    return;
end

if isKey(colorMap, char(colorText))
    out = colorMap(char(colorText));
    return;
end

knownColors = ["y","m","c","r","g","b","w","k", ...
    "yellow","magenta","cyan","red","green","blue","white","black"];
if any(colorText == knownColors)
    out = char(colorText);
end
end

function colorMap = buildColorMap(plotProperties)
colorMap = containers.Map('KeyType', 'char', 'ValueType', 'any');
if isempty(plotProperties)
    return;
end

plotVars = string(plotProperties.Properties.VariableNames);
if ~ismember("Plot Line Colours", plotVars) || ~ismember("Colour Hex Code", plotVars)
    return;
end

for iRow = 1:height(plotProperties)
    name = lower(strtrim(string(plotProperties.("Plot Line Colours")(iRow))));
    hexText = strtrim(string(plotProperties.("Colour Hex Code")(iRow)));
    if strlength(name) == 0 || ismissing(name) || strlength(hexText) == 0 || ismissing(hexText)
        continue;
    end

    rgb = hexToRgb(hexText);
    if ~isempty(rgb)
        colorMap(char(name)) = rgb;
    end
end
end

function rgb = hexToRgb(hexText)
rgb = [];
hexText = char(strtrim(string(hexText)));
hexText = strrep(hexText, '"', '');
hexText = strrep(hexText, '''', '');
if startsWith(hexText, '#')
    hexText = hexText(2:end);
end
if numel(hexText) ~= 6
    return;
end

try
    rgb = [hex2dec(hexText(1:2)) hex2dec(hexText(3:4)) hex2dec(hexText(5:6))] / 255;
catch
    rgb = [];
end
end

function [value, ok] = tryEvaluateEquation(eqText, context)
value = [];

ctxNames = fieldnames(context);
for iName = 1:numel(ctxNames)
    varName = ctxNames{iName};
    if isvarname(varName)
        eval([varName ' = context.(varName);']);
    end
end

try
    value = eval(eqText);
    ok = true;
catch
    ok = false;
end
end

function tf = isSupportedScalar(value)
if isnumeric(value) || islogical(value)
    tf = isscalar(value) && isfinite(double(value));
    return;
end

if isstring(value)
    tf = isscalar(value) && strlength(value) > 0;
    return;
end

if ischar(value)
    tf = ~isempty(value);
    return;
end

tf = false;
end

function out = formatValueWithUnit(value, unitText)
if isnumeric(value) || islogical(value)
    valueText = string(sprintf('%.2f', double(value)));
elseif isstring(value)
    valueText = value;
elseif ischar(value)
    valueText = string(value);
else
    out = "N.A";
    return;
end

unitText = strtrim(string(unitText));
if strlength(unitText) == 0 || ismissing(unitText)
    out = valueText;
else
    out = valueText + " " + unitText;
end
end

function out = formatValueOnly(value)
if isnumeric(value) || islogical(value)
    out = string(sprintf('%.2f', double(value)));
elseif isstring(value)
    out = value;
elseif ischar(value)
    out = string(value);
else
    out = "N.A";
end
end
