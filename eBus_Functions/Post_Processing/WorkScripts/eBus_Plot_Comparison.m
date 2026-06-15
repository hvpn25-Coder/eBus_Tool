function eBus_Plot_Comparison()
% eBus_Plot_Comparison  Compare Plot-bank figures from two MAT files.

thisScriptDir = fileparts(mfilename('fullpath'));
plotBankPath = fullfile(thisScriptDir, '..', 'KPIs_Plots', 'eBus_KPIs_Plots_Bank.xlsx');

state = struct();
state.matPaths = strings(1, 2);
state.matLabels = strings(1, 2);
state.contexts = cell(1, 2);
state.plotConfig = table();
state.groupNames = strings(0, 1);

if ~isfile(plotBankPath)
    errordlg(sprintf('KPI/Plot bank was not found:\n%s', plotBankPath), 'Missing Plot Bank');
    return;
end

try
    state.plotConfig = readtable(plotBankPath, ...
        'Sheet', 'Plots', ...
        'VariableNamingRule', 'preserve', ...
        'TextType', 'string');
    validatePlotConfig(state.plotConfig);
    state.groupNames = getPlotGroupNames(state.plotConfig);
catch ME
    errordlg(sprintf('Unable to read Plots sheet:\n%s', ME.message), 'Plot Bank Error');
    return;
end

hFig = figure( ...
    'Name', 'eBus Plot Comparison', ...
    'NumberTitle', 'off', ...
    'MenuBar', 'none', ...
    'ToolBar', 'none', ...
    'Resize', 'off', ...
    'Position', centerFigurePosition(760, 420), ...
    'CloseRequestFcn', @closeGui);

uicontrol(hFig, 'Style', 'text', ...
    'String', 'First MAT File', ...
    'HorizontalAlignment', 'left', ...
    'Position', [30 365 120 20]);
firstPathEdit = uicontrol(hFig, 'Style', 'edit', ...
    'String', '', ...
    'HorizontalAlignment', 'left', ...
    'Enable', 'inactive', ...
    'Position', [150 362 470 26]);
uicontrol(hFig, 'Style', 'pushbutton', ...
    'String', 'Select', ...
    'Position', [640 361 80 28], ...
    'Callback', @(~, ~)selectMatFile(1));

uicontrol(hFig, 'Style', 'text', ...
    'String', 'Second MAT File', ...
    'HorizontalAlignment', 'left', ...
    'Position', [30 325 120 20]);
secondPathEdit = uicontrol(hFig, 'Style', 'edit', ...
    'String', '', ...
    'HorizontalAlignment', 'left', ...
    'Enable', 'inactive', ...
    'Position', [150 322 470 26]);
uicontrol(hFig, 'Style', 'pushbutton', ...
    'String', 'Select', ...
    'Position', [640 321 80 28], ...
    'Callback', @(~, ~)selectMatFile(2));

groupLabel = uicontrol(hFig, 'Style', 'text', ...
    'String', 'Select Plot Group(s)', ...
    'HorizontalAlignment', 'left', ...
    'Visible', 'off', ...
    'Position', [30 275 250 20]);
groupList = uicontrol(hFig, 'Style', 'listbox', ...
    'String', cellstr(state.groupNames), ...
    'Min', 0, ...
    'Max', max(2, numel(state.groupNames)), ...
    'Value', [], ...
    'Visible', 'off', ...
    'Position', [30 90 690 180]);
plotButton = uicontrol(hFig, 'Style', 'pushbutton', ...
    'String', 'Plot Selected Groups', ...
    'Visible', 'off', ...
    'Position', [560 45 160 30], ...
    'Callback', @plotSelectedGroups);
statusText = uicontrol(hFig, 'Style', 'text', ...
    'String', 'Select two MAT files to continue.', ...
    'HorizontalAlignment', 'left', ...
    'Position', [30 45 500 30]);

    function selectMatFile(fileIdx)
        [fileName, filePath] = uigetfile('*.mat', sprintf('Select MAT file %d', fileIdx));
        if isequal(fileName, 0)
            return;
        end

        fullPath = string(fullfile(filePath, fileName));
        try
            context = load(fullPath);
        catch ME
            errordlg(sprintf('Unable to load MAT file:\n%s\n\n%s', fullPath, ME.message), 'MAT Load Error');
            return;
        end

        state.matPaths(fileIdx) = fullPath;
        [~, state.matLabels(fileIdx), ~] = fileparts(fullPath);
        state.contexts{fileIdx} = context;

        if fileIdx == 1
            set(firstPathEdit, 'String', char(fullPath));
        else
            set(secondPathEdit, 'String', char(fullPath));
        end

        if hasBothMatFiles()
            set(groupLabel, 'Visible', 'on');
            set(groupList, 'Visible', 'on');
            set(plotButton, 'Visible', 'on');
            set(statusText, 'String', 'Select one or more plot groups, then click Plot Selected Groups.');
        end
    end

    function tf = hasBothMatFiles()
        tf = all(strlength(state.matPaths) > 0) && all(cellfun(@isstruct, state.contexts));
    end

    function plotSelectedGroups(~, ~)
        selectedIdx = get(groupList, 'Value');
        if isempty(selectedIdx)
            errordlg('Select at least one plot group.', 'No Group Selected');
            return;
        end
        if ~hasBothMatFiles()
            errordlg('Select both MAT files before plotting.', 'MAT Files Required');
            return;
        end

        selectedGroups = state.groupNames(selectedIdx);
        set(statusText, 'String', 'Generating comparison plots...');
        drawnow;

        try
            plotComparisonGroups(state.plotConfig, state.contexts, state.matLabels, selectedGroups);
            set(statusText, 'String', 'Comparison plots generated.');
        catch ME
            set(statusText, 'String', 'Plot generation failed.');
            errordlg(sprintf('Unable to generate comparison plots:\n%s', ME.message), 'Plot Error');
        end
    end

    function closeGui(~, ~)
        if ishghandle(hFig)
            delete(hFig);
        end
    end
end

function plotComparisonGroups(plotConfig, contexts, fileLabels, selectedGroups)
allGroups = normalizeGroupValues(plotConfig.("Group"));
selectedGroups = string(selectedGroups(:));

for iGroup = 1:numel(selectedGroups)
    groupName = selectedGroups(iGroup);
    groupRows = plotConfig(allGroups == groupName, :);
    if isempty(groupRows)
        continue;
    end

    figureNos = unique(groupRows.("Figure No"), 'stable');
    for iFig = 1:numel(figureNos)
        figNo = figureNos(iFig);
        figRows = groupRows(groupRows.("Figure No") == figNo, :);
        if ~hasYRows(figRows)
            continue;
        end

        subplotNos = unique(figRows.("Subplot"), 'stable');
        subplotNos = subplotNos(~isnan(double(subplotNos)));
        if isempty(subplotNos)
            continue;
        end

        figTitle = sprintf('Figure %s _ %s _ Comparison', char(string(figNo)), char(groupName));
        hPlot = figure('Name', figTitle, 'NumberTitle', 'off', 'Color', 'w');
        tileObj = tiledlayout(hPlot, numel(subplotNos), 1, 'TileSpacing', 'compact', 'Padding', 'compact');
        title(tileObj, figTitle, 'Interpreter', 'none');
        axesList = gobjects(numel(subplotNos), 1);

        for iSub = 1:numel(subplotNos)
            ax = nexttile(tileObj, iSub);
            axesList(iSub) = ax;
            plotComparisonSubplot(ax, figRows(figRows.("Subplot") == subplotNos(iSub), :), contexts, fileLabels);
        end

        linkSubplotXAxis(axesList);
    end
end
end

function plotComparisonSubplot(ax, subRows, contexts, fileLabels)
hold(ax, 'on');
grid(ax, 'on');

axisKinds = upper(strtrim(string(subRows.("Axis"))));
titleRows = subRows(axisKinds == "T", :);
if ~isempty(titleRows)
    subTitle = strtrim(string(titleRows.("Signals and Titles")(1)));
    if strlength(subTitle) > 0 && ~ismissing(subTitle)
        title(ax, subTitle, 'Interpreter', 'none');
    end
end

xRows = subRows(axisKinds == "X", :);
yRows = subRows(axisKinds == "Y", :);
if isempty(yRows)
    hold(ax, 'off');
    return;
end

xDataByFile = cell(1, 2);
hasXData = false(1, 2);
for iFile = 1:2
    [xDataByFile{iFile}, hasXData(iFile)] = resolveXAxisData(xRows, contexts{iFile});
end

if ~isempty(xRows)
    xLabelText = buildAxisLabel(xRows.("Label")(1), xRows.("Unit")(1));
    if strlength(xLabelText) > 0
        xlabel(ax, xLabelText, 'Interpreter', 'none');
    end
end

colorOrder = get(ax, 'ColorOrder');
if isempty(colorOrder)
    colorOrder = lines(max(1, height(yRows)));
end

hasLeftYLabel = false;
hasRightYLabel = false;
for iY = 1:height(yRows)
    yExpr = strtrim(string(yRows.("Signals and Titles")(iY)));
    if strlength(yExpr) == 0 || ismissing(yExpr)
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

    lineColor = colorOrder(mod(iY - 1, size(colorOrder, 1)) + 1, :);
    lineWidth = resolveLineWidth(yRows.("Width")(iY));
    legendBase = resolveLegendText(yRows, iY, yExpr);

    for iFile = 1:2
        [yData, okY] = tryEvaluateEquation(yExpr, contexts{iFile});
        if ~okY || ~isnumeric(yData) || isempty(yData)
            continue;
        end

        styleArg = "-";
        if iFile == 2
            styleArg = "--";
        end

        if hasXData(iFile) && numel(xDataByFile{iFile}) == numel(yData)
            p = plot(ax, xDataByFile{iFile}, yData, styleArg, ...
                'LineWidth', lineWidth, 'Color', lineColor);
        else
            p = plot(ax, yData, styleArg, ...
                'LineWidth', lineWidth, 'Color', lineColor);
        end
        p.DisplayName = sprintf('%s - %s', char(fileLabels(iFile)), char(legendBase));
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

legend(ax, 'show', 'Location', 'best', 'Interpreter', 'none');
hold(ax, 'off');
end

function [xData, hasData] = resolveXAxisData(xRows, context)
xData = [];
hasData = false;
if isempty(xRows)
    return;
end

xExpr = strtrim(string(xRows.("Signals and Titles")(1)));
if strlength(xExpr) == 0 || ismissing(xExpr)
    return;
end

[xCandidate, okX] = tryEvaluateEquation(xExpr, context);
if okX && isnumeric(xCandidate) && ~isempty(xCandidate)
    xData = xCandidate;
    hasData = true;
end
end

function legendText = resolveLegendText(yRows, rowIdx, fallbackText)
legendText = strtrim(string(fallbackText));
if ismember("Legend", string(yRows.Properties.VariableNames))
    candidate = strtrim(string(yRows.("Legend")(rowIdx)));
    if strlength(candidate) > 0 && ~ismissing(candidate)
        legendText = candidate;
    end
end
end

function validatePlotConfig(plotConfig)
requiredCols = ["Group", "Figure No", "Subplot", "Axis", "Signals and Titles"];
for iCol = 1:numel(requiredCols)
    if ~ismember(requiredCols(iCol), string(plotConfig.Properties.VariableNames))
        error('Required column "%s" was not found in sheet "Plots".', requiredCols(iCol));
    end
end
end

function groupNames = getPlotGroupNames(plotConfig)
if isempty(plotConfig)
    groupNames = strings(0, 1);
    return;
end
groupNames = unique(normalizeGroupValues(plotConfig.("Group")), 'stable');
groupNames = groupNames(:);
end

function out = normalizeGroupValues(groupCol)
out = strtrim(string(groupCol));
out(ismissing(out) | strlength(out) == 0) = "GENERAL";
end

function tf = hasYRows(figRows)
tf = false;
if isempty(figRows) || ~ismember("Axis", string(figRows.Properties.VariableNames))
    return;
end
axisKinds = upper(strtrim(string(figRows.("Axis"))));
tf = any(axisKinds == "Y");
end

function linkSubplotXAxis(axesList)
validAxes = axesList(isgraphics(axesList, 'axes'));
if numel(validAxes) <= 1
    return;
end
try
    linkaxes(validAxes, 'x');
catch
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
    value = eval(char(string(eqText)));
    ok = true;
catch
    ok = false;
end
end

function pos = centerFigurePosition(width, height)
screenSize = get(0, 'ScreenSize');
left = max(1, round((screenSize(3) - width) / 2));
bottom = max(1, round((screenSize(4) - height) / 2));
pos = [left bottom width height];
end
