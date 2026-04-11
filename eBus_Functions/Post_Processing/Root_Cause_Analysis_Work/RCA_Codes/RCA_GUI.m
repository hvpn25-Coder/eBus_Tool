function fig = RCA_GUI(resultsInput)
% RCA_GUI  MATLAB-native interactive viewer for eBus RCA results.
%
% Usage:
%   RCA_GUI(RCA_Results)
%   RCA_GUI('C:\path\to\RCA_Results.mat')
%   RCA_GUI
%
% The GUI follows the same 18-section flow as the Word report. The right
% pane is scrollable and contains text, tables, and figure preview blocks.
% Each preview can be opened as a normal MATLAB figure with zoom enabled.

if nargin < 1
    resultsInput = [];
end

[results, sourceInfo] = localResolveResults(resultsInput);
sections = localBuildSections(results, sourceInfo);
items = cellstr([sections.NavTitle]');

fig = uifigure('Name', 'eBus RCA Interactive Viewer', ...
    'Position', [80 60 1500 860], ...
    'Color', [0.94 0.96 0.98]);

mainGrid = uigridlayout(fig, [1 2]);
mainGrid.ColumnWidth = {285, '1x'};
mainGrid.Padding = [14 14 14 14];
mainGrid.ColumnSpacing = 14;

navPanel = uipanel(mainGrid, 'BackgroundColor', [0.06 0.13 0.20], 'BorderType', 'none');
navPanel.Layout.Column = 1;
navGrid = uigridlayout(navPanel, [5 1]);
navGrid.RowHeight = {58, 30, '1x', 34, 34};
navGrid.Padding = [12 14 12 14];
navGrid.RowSpacing = 10;

uilabel(navGrid, 'Text', 'eBus RCA', 'FontSize', 22, 'FontWeight', 'bold', ...
    'FontColor', [0.94 0.98 1.00], 'HorizontalAlignment', 'left');
uilabel(navGrid, 'Text', 'Sections 1 to 18', 'FontSize', 12, 'FontWeight', 'bold', ...
    'FontColor', [0.66 0.78 0.88], 'HorizontalAlignment', 'left');
sectionList = uilistbox(navGrid, 'Items', items, 'Value', items{1}, ...
    'FontName', 'Calibri', 'FontSize', 13, 'BackgroundColor', [0.91 0.95 0.98]);
uibutton(navGrid, 'Text', 'Open Output Folder', 'FontWeight', 'bold', ...
    'ButtonPushedFcn', @(~, ~) localOpenFolder(localOutputFolder(results)));
refreshButton = uibutton(navGrid, 'Text', 'Refresh Section', 'FontWeight', 'bold');

contentOuter = uipanel(mainGrid, 'BackgroundColor', [0.94 0.96 0.98], 'BorderType', 'none');
contentOuter.Layout.Column = 2;
contentGrid = uigridlayout(contentOuter, [2 1]);
contentGrid.RowHeight = {86, '1x'};
contentGrid.Padding = [0 0 0 0];
contentGrid.RowSpacing = 12;

headerPanel = uipanel(contentGrid, 'BackgroundColor', [1 1 1], 'BorderType', 'none');
headerGrid = uigridlayout(headerPanel, [2 3]);
headerGrid.RowHeight = {40, 26};
headerGrid.ColumnWidth = {'1x', 210, 170};
headerGrid.Padding = [18 10 18 8];
titleLabel = uilabel(headerGrid, 'Text', sections(1).Title, 'FontSize', 23, ...
    'FontWeight', 'bold', 'FontColor', [0.07 0.18 0.29]);
titleLabel.Layout.Row = 1;
titleLabel.Layout.Column = 1;
sourceLabel = uilabel(headerGrid, 'Text', localShortSource(sourceInfo), ...
    'FontSize', 11, 'FontColor', [0.35 0.40 0.46]);
sourceLabel.Layout.Row = 2;
sourceLabel.Layout.Column = 1;
enButton = uibutton(headerGrid, 'Text', 'Generate Word Report', 'FontWeight', 'bold', ...
    'ButtonPushedFcn', @(~, ~) localGenerateReport(results, 'EN'));
enButton.Layout.Row = [1 2];
enButton.Layout.Column = 2;
deButton = uibutton(headerGrid, 'Text', 'German Report', 'FontWeight', 'bold', ...
    'ButtonPushedFcn', @(~, ~) localGenerateReport(results, 'DE'));
deButton.Layout.Row = [1 2];
deButton.Layout.Column = 3;

contentPanel = uipanel(contentGrid, 'BackgroundColor', [0.94 0.96 0.98], 'BorderType', 'none');
try
    contentPanel.Scrollable = 'on';
    contentPanel.AutoResizeChildren = 'off';
catch
end

sectionList.ValueChangedFcn = @(src, ~) localRenderFromList(src, contentPanel, titleLabel, sourceLabel, sections, sourceInfo);
refreshButton.ButtonPushedFcn = @(~, ~) localRenderFromList(sectionList, contentPanel, titleLabel, sourceLabel, sections, sourceInfo);
drawnow;
localRenderSection(contentPanel, titleLabel, sourceLabel, sections(1), sourceInfo);
end

function [results, sourceInfo] = localResolveResults(resultsInput)
results = [];
sourceInfo = struct('SourceType', "None", 'SourcePath', "", 'Description', "No RCA results supplied.");
if isempty(resultsInput)
    try
        resultsInput = evalin('base', 'RCA_Results');
    catch
        resultsInput = localLatestResultsFile();
    end
end
if isstruct(resultsInput)
    results = resultsInput;
    sourceInfo.SourceType = "Workspace struct";
    sourceInfo.Description = "RCA results supplied from MATLAB workspace.";
    sourceInfo.SourcePath = localMatFile(results);
    return;
end
if ischar(resultsInput) || (isstring(resultsInput) && isscalar(resultsInput))
    filePath = char(string(resultsInput));
    if ~isfile(filePath)
        error('RCA_GUI:MissingResults', 'RCA results file not found: %s', filePath);
    end
    loaded = load(filePath);
    names = fieldnames(loaded);
    for iName = 1:numel(names)
        candidate = loaded.(names{iName});
        if isstruct(candidate) && (isfield(candidate, 'VehicleKPI') || isfield(candidate, 'SubsystemResults'))
            results = candidate;
            break;
        end
    end
    if isempty(results)
        error('RCA_GUI:InvalidResults', 'No RCA results struct was found in %s.', filePath);
    end
    sourceInfo.SourceType = "MAT file";
    sourceInfo.Description = "RCA results loaded from MAT file.";
    sourceInfo.SourcePath = string(filePath);
    return;
end
error('RCA_GUI:UnsupportedInput', 'Input must be an RCA results struct or RCA_Results.mat path.');
end

function filePath = localLatestResultsFile()
rootFolder = fileparts(fileparts(mfilename('fullpath')));
matches = dir(fullfile(rootFolder, '**', 'RCA_Results.mat'));
if isempty(matches)
    filePath = [];
else
    [~, idx] = max([matches.datenum]);
    filePath = fullfile(matches(idx).folder, matches(idx).name);
end
end

function sections = localBuildSections(results, sourceInfo)
titles = ["1. Info"; "2. Document Control"; "3. Table of Contents"; ...
    "4. List of Figures"; "5. List of Tables"; "6. Abbreviations / Nomenclature"; ...
    "7. Introduction"; "8. Simulation and Data Overview"; "9. Technical Summary"; ...
    "10. Analysis Methodology"; "11. Vehicle-Level Assessment"; "12. Subsystem-Level RCA"; ...
    "13. Event-Based Deep Dives"; "14. KPI Dashboard Summary"; "15. Root Cause Summary Table"; ...
    "16. Recommendations"; "17. Conclusion"; "18. Appendices"];
sections = repmat(struct('Title', "", 'NavTitle', "", 'Blocks', localBlock("text", "", "", table(), strings(0, 1))), numel(titles), 1);
for iSection = 1:numel(titles)
    sections(iSection).Title = titles(iSection);
    sections(iSection).NavTitle = titles(iSection);
end
sections(1).Blocks = [localText('How to use this GUI', localJoin(["Use the left navigation pane to move through Sections 1 to 18 in the same order as the Word report."; "The right pane is scrollable and contains report-style text, KPI tables, RCA tables, and figure evidence."; "Figure cards include Open Interactive Figure, which opens the saved plot in a normal MATLAB figure window with zoom enabled."])); localTable('Run and report context', localRunContext(results, sourceInfo))];
sections(2).Blocks = [localText('Document purpose', localJoin(["This RCA package converts electric bus simulation outputs into a formal technical evidence set for engineering review and decision-making."; "It links vehicle-level behaviour, subsystem contributors, event-based evidence, and recommended actions."; "The GUI uses the same result structure as the Word report, allowing fast MATLAB-native review before optional document generation."])); localText('Scope', localJoin(["The scope covers the supplied simulation data set and available workbook metadata."; "The workflow includes signal audit, KPI generation, event segmentation, bad-segment selection, root-cause ranking, subsystem drill-down, and recommendations."; "Missing signals are limitations, not execution blockers. Approximate conclusions must be validated through model review, calibration review, or test correlation."])); localTable('Version history', localVersionHistory()); localTable('Review and approval', localReviewTable()); localTable('Analysis output inventory', localTableInventory(results))];
sections(3).Blocks = localTable('Full table of contents', localTocTable(titles));
sections(4).Blocks = [localText('Figure list', "Saved RCA figure files found in the result structure."); localTable('Available RCA figures', localFigureList(localAllFigures(results))); localFigure('All RCA figures', localAllFigures(results))];
sections(5).Blocks = localTable('Available RCA tables', localTableInventory(results));
sections(6).Blocks = localTable('Abbreviations and nomenclature', localAbbreviations());
sections(7).Blocks = [localText('Background and objective', localJoin(["The eBus simulation program evaluates route energy, traction performance, control behaviour, subsystem losses, and operating sensitivity before physical test exposure."; "The RCA objective is to identify major efficiency, performance, operation, and energy issues and connect them to likely physical, control, calibration, or model-quality contributors."])); localText('Questions answered', localJoin(["What did the bus do over the drive cycle and where did it deviate from expected behaviour?"; "Which segments consume excessive energy, show poor tracking, or exhibit high losses?"; "Which subsystems likely explain the symptoms and what should be changed next?"]))];
sections(8).Blocks = [localText('Model overview', localJoin(["The Excel workbook is the primary source for subsystem descriptions, signal names, units, sign conventions, and specification evaluation logic."; "The MAT file is the logged numerical evidence set. Signal evaluation fallbacks and extraction heuristics handle practical logging variations."; "Module overview and simulation configuration content are shown when available in the RCA result context."])); localTable('Data source and loading context', localRunContext(results, sourceInfo)); localTable('Simulation configuration overview', localSimulationConfig(results)); localTable('MAT inventory snapshot', localLimit(localGetTable(results, 'MatInventory'), 40)); localTable('Detailed signal list and presence status', localLimit(localTrimSignalTable(localGetTable(results, 'SignalPresence')), 80)); localTable('Specification presence status', localLimit(localGetTable(results, 'SpecPresence'), 80)); localFigure('Model and data overview figures', localFindFigures(results, ["module", "model", "simulation", "overview", "availability"]))];
sections(9).Blocks = [localText('Why this analysis was performed', localJoin(["The analysis converts simulation logs into management-ready and subsystem-owner-ready RCA evidence."; "Symptoms are quantified, likely contributors are ranked, and actions are tied to available signals."])); localText('Vehicle narrative', localTextField(results, 'VehicleNarrative', "Vehicle-level narrative was not recorded.")); localText('Root-cause narrative', localTextField(results, 'RootCauseNarrative', "Root-cause narrative was not recorded.")); localTable('Top vehicle KPIs', localLimit(localGetTable(results, 'VehicleKPI'), 16)); localTable('Top root-cause ranking', localLimit(localGetTable(results, 'RootCauseRanking'), 12)); localTable('Recommended actions summary', localLimit(localGetTable(results, 'OptimizationTable'), 12)); localFigure('Executive figures', localFindFigures(results, ["energy_flow", "segment_ranking", "root", "pareto"]))];
sections(10).Blocks = [localText('Overall RCA workflow', localJoin(["1. Read workbook metadata and MAT-file inventory."; "2. Resolve signals and specifications using workbook fallbacks."; "3. Extract traces, align them to common time, and repair duplicate/invalid samples."; "4. Build derived speed, acceleration, energy, power, force, loss, and operating-class traces."; "5. Segment the drive cycle into event/context intervals."; "6. Compute vehicle, segment, and subsystem KPIs."; "7. Detect bad segments using thresholds and severity ranking."; "8. Score root-cause candidates, normalize contribution shares, and assign confidence."; "9. Generate subsystem feedback, vehicle conclusions, GUI content, and optional Word reports."])); localText('Event-based segmentation', localJoin(["Segments are created from changes in motion class, grade class, auxiliary class, dominant gear, stop/start events, and minimum-duration constraints."; "Adjacent segments with the same dominant classes are merged to reduce artificial fragmentation."; "SegmentSummary drives segment KPI calculation, bad-segment detection, and segment-by-segment RCA."])); localText('Bad segment and root-cause logic', localJoin(["Bad segments are selected from explicit flags or severity scores based on energy intensity, tracking error, loss share, regen effectiveness, auxiliary burden, and subsystem evidence."; "Candidate causes receive evidence scores from physical factors such as controller tracking, gear operation, electrical limitation, transmission loss, road load, grade, braking split, and available power/force margins."; "Contribution percentage is each cause score normalized by the sum of positive scores. The primary cause is the highest contribution cause. Confidence increases with independent supporting signals and decreases when key signals are missing."])); localTable('Configured thresholds and assumptions', localLimit(localThresholdTable(), 80)); localFigure('Methodology and balance figures', localFindFigures(results, ["workflow", "method", "balance", "force", "power", "segment"]))];
sections(11).Blocks = localVehicleBlocks(results);
sections(12).Blocks = localSubsystemBlocks(results);
sections(13).Blocks = [localText('Event classes', localJoin(["Acceleration events: response delay, pedal saturation, torque availability, gear choice, grade burden, tracking shortfall."; "Braking events: friction/regeneration split, recuperation opportunity, pedal behavior, deceleration consistency."; "Cruising events: speed-hold quality, auxiliary burden, road load, gear stability, unnecessary torque transients."; "Hill-climb and low-SOC events: power margin, current limitation, voltage sag, derating, and performance penalty."])); localTable('Segment summary', localLimit(localGetTable(results, 'SegmentSummary'), 100)); localTable('Bad segment table', localLimit(localGetTable(results, 'BadSegmentTable'), 100)); localFigure('Event and worst-segment figures', localFindFigures(results, ["event", "worst", "bad_segment", "driver", "segment"]))];
sections(14).Blocks = [localText('KPI dashboard purpose', "This dashboard collects the main vehicle, segment, subsystem, and data-quality KPI tables used by the report."); localTable('Vehicle KPI dashboard', localLimit(localGetTable(results, 'VehicleKPI'), 160)); localTable('Segment KPI dashboard', localLimit(localGetTable(results, 'SegmentSummary'), 160)); localTable('Subsystem KPI dashboard', localLimit(localSubsystemKpis(results), 200)); localTable('Signal availability dashboard', localLimit(localTrimSignalTable(localGetTable(results, 'SignalPresence')), 160))];
sections(15).Blocks = [localText('Root-cause summary interpretation', localJoin(["The root-cause table is the high-value action table for engineering review."; "A high contribution percentage means the segment evidence is most consistent with that contributor among evaluated candidates."; "Recommended ownership should be assigned after reviewing signal basis, confidence note, and subsystem plots."])); localTable('Root-cause ranking table', localLimit(localGetTable(results, 'RootCauseRanking'), 160)); localTable('Bad-segment RCA table', localLimit(localGetTable(results, 'BadSegmentTable'), 160)); localFigure('Root-cause summary figures', localFindFigures(results, ["root", "pareto", "ranking", "bad_segment"]))];
sections(16).Blocks = [localText('Recommendation logic', localJoin(["Recommendations are split into immediate actions, model improvements, controls/calibration improvements, design opportunities, and validation work."; "The highest-value actions address repeated bad-segment patterns or large energy/loss contributors with high-confidence evidence."; "Additional logging is recommended wherever an RCA conclusion remains approximate because a key internal signal is missing."])); localTable('Recommended actions', localLimit(localGetTable(results, 'OptimizationTable'), 160)); localTable('Subsystem-owner action hints', localLimit(localSubsystemSuggestions(results), 160))];
sections(17).Blocks = [localText('Engineering conclusion', localJoin(["The RCA is an evidence-based prioritization of vehicle and subsystem improvement opportunities."; "The main limitations are summarized by the vehicle KPI table, bad-segment table, and root-cause ranking."; "The next step is to assign owners to top causes, review corresponding subsystem plots, and rerun RCA after calibration or model changes."])); localText('Recorded vehicle narrative', localTextField(results, 'VehicleNarrative', "No vehicle conclusion narrative was recorded.")); localText('Recorded root-cause narrative', localTextField(results, 'RootCauseNarrative', "No root-cause conclusion narrative was recorded."))];
sections(18).Blocks = [localText('Appendix usage', "Appendices provide detailed signal lists, extraction logs, full KPI tables, and traceability to MATLAB scripts."); localTable('Detailed signal list and presence status', localLimit(localTrimSignalTable(localGetTable(results, 'SignalPresence')), 240)); localTable('Specification list and presence status', localLimit(localGetTable(results, 'SpecPresence'), 180)); localTable('Signal extraction log', localLimit(localGetTable(results, 'ExtractionLog'), 180)); localTable('Full vehicle KPI table', localLimit(localGetTable(results, 'VehicleKPI'), 240)); localTable('Full segment summary table', localLimit(localGetTable(results, 'SegmentSummary'), 240)); localTable('MATLAB script references', localScriptRefs())];
end

function blocks = localVehicleBlocks(results)
kpi = localGetTable(results, 'VehicleKPI');
blocks = [localText('Vehicle speed tracking', localJoin(["This subsection reviews actual speed versus desired speed, tracking error, route grade, acceleration, and gear usage."; "Tracking errors are interpreted with drive context. Uphill acceleration error may indicate torque, power, gear, or controller limitation; cruise error may indicate calibration or feedforward mismatch."; "Distance error between target and actual route completion is treated as a cumulative vehicle performance KPI."])); localTable('Vehicle speed tracking and performance KPI summary', localFilterKpi(kpi, ["speed", "tracking", "performance", "distance", "accel"])); localFigure('Vehicle speed tracking over drive cycle', localFindFigures(results, ["speed_tracking", "vehicle_speed", "tracking", "performance"])); localText('Energy, efficiency, and range impact', localJoin(["Energy consumption is evaluated using battery energy, traction energy, auxiliary energy, braking/regen energy, and losses where signals permit."; "Energy consumption per kilometre directly links trip behaviour to range and operating cost."; "Negative segment energy intensity can occur during net recuperation, so loss and auxiliary shares must be interpreted against discharge-domain energy rather than signed net segment energy."])); localTable('Energy, efficiency, and range KPI summary', localFilterKpi(kpi, ["energy", "efficiency", "range", "Wh/km", "loss", "aux", "regen"])); localFigure('Vehicle energy, loss, and range figures', localFindFigures(results, ["energy", "efficiency", "range", "loss", "aux", "regen"])); localText('Power balance and force balance', localJoin(["Power balance compares battery, DC bus, motor, transmission, wheel, auxiliary, loss, and braking domains."; "Force balance compares tractive force, braking force, aerodynamic drag, rolling resistance, gradient force, and acceleration response."; "Large unexplained residuals indicate missing signals, sign-convention issues, model-energy inconsistency, or post-processing assumptions requiring review."])); localTable('Vehicle-level KPI table', localLimit(kpi, 100)); localFigure('Power and force balance figures', localFindFigures(results, ["power_balance", "force_balance", "balance", "energy_flow"]))];
end

function blocks = localSubsystemBlocks(results)
subs = localSubsystems(results);
blocks = localText('Subsystem RCA overview', localJoin(["Subsystem content is populated from RCA_Results.SubsystemResults."; "Each subsystem is shown with interpretation text, KPI table, warnings/recommendations when recorded, and saved figures."]));
if isempty(subs)
    blocks(end + 1) = localText('Subsystem RCA availability', "No subsystem RCA results were recorded.");
    return;
end
for iSub = 1:numel(subs)
    subName = localSubsystemName(subs(iSub), iSub);
    blocks(end + 1) = localText(subName + " interpretation", localSubsystemText(subs(iSub))); %#ok<AGROW>
    blocks(end + 1) = localTable(subName + " KPI table", localLimit(localSubsystemKpi(subs(iSub)), 100)); %#ok<AGROW>
    blocks(end + 1) = localFigure(subName + " figures", localSubsystemFigures(subs(iSub))); %#ok<AGROW>
end
end

function localRenderFromList(listBox, contentPanel, titleLabel, sourceLabel, sections, sourceInfo)
items = listBox.Items;
idx = find(strcmp(items, listBox.Value), 1, 'first');
if isempty(idx)
    idx = 1;
end
localRenderSection(contentPanel, titleLabel, sourceLabel, sections(idx), sourceInfo);
end

function localRenderSection(panel, titleLabel, sourceLabel, section, sourceInfo)
delete(panel.Children);
titleLabel.Text = section.Title;
sourceLabel.Text = sprintf('%s | %d content blocks', localShortSource(sourceInfo), numel(section.Blocks));
drawnow limitrate;
pos = panel.Position;
if numel(pos) < 4 || pos(3) < 100 || pos(4) < 100
    pos = [0 0 1120 720];
end
w = max(780, pos(3) - 34);
x = 16;
gap = 14;
blockHeights = zeros(numel(section.Blocks), 1);
for iBlock = 1:numel(section.Blocks)
    blockHeights(iBlock) = localBlockHeight(section.Blocks(iBlock), w);
end
totalH = max(pos(4) + 40, 84 + sum(blockHeights) + gap * max(numel(section.Blocks) - 1, 0));
y = totalH - 62;
uilabel(panel, 'Text', section.Title, 'Position', [x y w 34], ...
    'FontSize', 24, 'FontWeight', 'bold', 'FontColor', [0.05 0.17 0.28], ...
    'BackgroundColor', [0.94 0.96 0.98]);
y = y - 44;
for iBlock = 1:numel(section.Blocks)
    h = blockHeights(iBlock);
    y = y - h;
    block = section.Blocks(iBlock);
    if block.Type == "text"
        localTextCard(panel, block, [x y w h]);
    elseif block.Type == "table"
        localTableCard(panel, block, [x y w h]);
    else
        localFigureCard(panel, block, [x y w h]);
    end
    y = y - gap;
end
end

function h = localBlockHeight(block, w)
if block.Type == "text"
    lineCount = max(3, numel(splitlines(string(block.Text))));
    h = min(285, max(130, 92 + 20 * lineCount));
elseif block.Type == "table"
    rowCount = 0;
    if istable(block.TableData)
        rowCount = height(block.TableData);
    end
    h = min(640, max(190, 116 + 27 * min(rowCount, 18)));
elseif isempty(block.Figures)
    h = 150;
else
    h = min(590, max(430, round(w * 0.36)));
end
end

function localTextCard(parent, block, p)
card = uipanel(parent, 'Position', p, 'BackgroundColor', [1 1 1], ...
    'BorderType', 'line', 'HighlightColor', [0.80 0.86 0.91]);
uilabel(card, 'Text', char(block.Title), 'Position', [16 p(4)-42 p(3)-32 26], ...
    'FontSize', 15, 'FontWeight', 'bold', 'FontColor', [0.04 0.20 0.33]);
ta = uitextarea(card, 'Value', cellstr(splitlines(string(block.Text))), ...
    'Position', [16 14 p(3)-32 p(4)-62], 'Editable', 'off', ...
    'FontName', 'Calibri', 'FontSize', 13, 'FontColor', [0.14 0.17 0.22], ...
    'BackgroundColor', [0.98 0.99 1.00]);
try
    ta.WordWrap = 'on';
catch
end
end

function localTableCard(parent, block, p)
card = uipanel(parent, 'Position', p, 'BackgroundColor', [1 1 1], ...
    'BorderType', 'line', 'HighlightColor', [0.80 0.86 0.91]);
tbl = localLimit(block.TableData, 250);
[data, cols] = localTableForUi(tbl);
uilabel(card, 'Text', sprintf('%s (%d rows shown)', char(block.Title), size(data, 1)), ...
    'Position', [16 p(4)-42 p(3)-32 26], 'FontSize', 15, 'FontWeight', 'bold', ...
    'FontColor', [0.04 0.20 0.33]);
if isempty(data)
    uitextarea(card, 'Value', {'No table rows are available for this evidence block.'}, ...
        'Position', [16 16 p(3)-32 p(4)-70], 'Editable', 'off', 'FontSize', 13, ...
        'BackgroundColor', [0.98 0.99 1.00]);
else
    uit = uitable(card, 'Data', data, 'ColumnName', cols, 'RowName', {}, ...
        'Position', [16 16 p(3)-32 p(4)-70], 'FontName', 'Calibri', 'FontSize', 12);
    localSetWidths(uit, data, cols, p(3)-32);
end
end

function localFigureCard(parent, block, p)
card = uipanel(parent, 'Position', p, 'BackgroundColor', [1 1 1], ...
    'BorderType', 'line', 'HighlightColor', [0.80 0.86 0.91]);
uilabel(card, 'Text', char(block.Title), 'Position', [16 p(4)-42 p(3)-32 26], ...
    'FontSize', 15, 'FontWeight', 'bold', 'FontColor', [0.04 0.20 0.33]);
files = localFigureFiles(block.Figures);
if isempty(files)
    uitextarea(card, 'Value', {'No saved figures were found for this section.'}, ...
        'Position', [16 16 p(3)-32 p(4)-70], 'Editable', 'off', 'FontSize', 13, ...
        'BackgroundColor', [0.98 0.99 1.00]);
    return;
end
items = cellstr(files);
dd = uidropdown(card, 'Items', items, 'Value', items{1}, ...
    'Position', [16 p(4)-78 p(3)-390 28], 'FontName', 'Calibri', 'FontSize', 12);
uibutton(card, 'Text', 'Open Interactive Figure', 'Position', [p(3)-360 p(4)-78 164 28], ...
    'FontWeight', 'bold', 'ButtonPushedFcn', @(~, ~) localOpenFigure(dd.Value));
uibutton(card, 'Text', 'Open File', 'Position', [p(3)-186 p(4)-78 78 28], ...
    'ButtonPushedFcn', @(~, ~) localOpenFile(dd.Value));
uibutton(card, 'Text', 'Folder', 'Position', [p(3)-98 p(4)-78 82 28], ...
    'ButtonPushedFcn', @(~, ~) localOpenFolder(fileparts(dd.Value)));
ax = uiaxes(card, 'Position', [16 16 p(3)-32 p(4)-108], 'Box', 'on', 'XTick', [], 'YTick', []);
try
    ax.Toolbar.Visible = 'on';
catch
end
localPreview(ax, dd.Value);
dd.ValueChangedFcn = @(src, ~) localPreview(ax, src.Value);
end

function b = localBlock(type, titleText, textValue, tableData, figures)
b = struct('Type', string(type), 'Title', string(titleText), 'Text', string(textValue), ...
    'TableData', tableData, 'Figures', localFigureFiles(figures));
end

function b = localText(titleText, textValue)
b = localBlock("text", titleText, textValue, table(), strings(0, 1));
end

function b = localTable(titleText, tableData)
if ~istable(tableData)
    tableData = table();
end
b = localBlock("table", titleText, "", tableData, strings(0, 1));
end

function b = localFigure(titleText, figures)
b = localBlock("figure", titleText, "", table(), figures);
end

function [data, cols] = localTableForUi(tbl)
data = {};
cols = {};
if ~istable(tbl) || height(tbl) == 0 || width(tbl) == 0
    return;
end
cols = tbl.Properties.VariableNames;
raw = table2cell(tbl);
data = cell(size(raw));
for iCell = 1:numel(raw)
    data{iCell} = char(localTextValue(raw{iCell}));
end
end

function localSetWidths(uit, data, cols, availableWidth)
try
    n = max(1, numel(cols));
    widths = repmat(max(100, floor(double(availableWidth) / n)), 1, n);
    for iCol = 1:n
        maxChars = strlength(string(cols{iCol}));
        if ~isempty(data)
            maxChars = max(maxChars, max(strlength(string(data(:, iCol)))));
        end
        widths(iCol) = min(360, max(90, 8 * double(min(maxChars, 45)) + 34));
    end
    uit.ColumnWidth = num2cell(widths);
catch
end
end

function tbl = localLimit(tbl, maxRows)
if nargin < 2
    maxRows = 100;
end
if ~istable(tbl)
    tbl = table();
    return;
end
if height(tbl) > maxRows
    tbl = tbl(1:maxRows, :);
end
for iVar = 1:width(tbl)
    if isstring(tbl.(iVar)) || iscategorical(tbl.(iVar))
        tbl.(iVar) = cellstr(string(tbl.(iVar)));
    end
end
end

function tbl = localGetTable(results, fieldName)
tbl = table();
if isstruct(results) && isfield(results, fieldName) && istable(results.(fieldName))
    tbl = results.(fieldName);
end
end

function tbl = localTrimSignalTable(tbl)
if ~istable(tbl)
    return;
end
removeNames = {'Note', 'Requirement', 'Method', 'BestMatch'};
keep = true(1, width(tbl));
for iName = 1:numel(removeNames)
    keep = keep & ~strcmpi(tbl.Properties.VariableNames, removeNames{iName});
end
tbl = tbl(:, keep);
end

function tbl = localRunContext(results, sourceInfo)
rows = { ...
    'RCA results source', char(sourceInfo.SourceType), char(sourceInfo.Description); ...
    'RCA results path', char(sourceInfo.SourcePath), 'Used to populate GUI content'; ...
    'Simulation MAT file', char(localMatFile(results)), 'Logged numerical evidence'; ...
    'Excel metadata file', char(localExcelFile(results)), 'Signal, specification, subsystem, and sign-convention source'; ...
    'Output folder', char(localOutputFolder(results)), 'Tables, PNG figures, MAT output, and optional Word reports'};
tbl = cell2table(rows, 'VariableNames', {'Item', 'Value', 'EngineeringNote'});
end

function tbl = localVersionHistory()
d = char(string(datetime('now', 'Format', 'dd-MMM-yyyy')));
rows = { ...
    '1.0', d, 'Initial RCA report and GUI structure', char(localAuthor()); ...
    '1.1', d, 'Added subsystem RCA, segment RCA, and report-aligned GUI evidence view', char(localAuthor()); ...
    '1.2', d, 'Updated model overview, simulation configuration, balance analysis, and interactive navigation', char(localAuthor())};
tbl = cell2table(rows, 'VariableNames', {'Version', 'Date', 'ChangeSummary', 'Author'});
end

function tbl = localReviewTable()
d = char(string(datetime('now', 'Format', 'dd-MMM-yyyy')));
rows = { ...
    'Simulation Engineer', char(localAuthor()), d, 'Prepared RCA and reviewed generated evidence'; ...
    'Simulation Lead', '[Insert Name]', '[Insert Date]', 'Review technical correctness and priority ranking'; ...
    'Module Owner', '[Insert Name]', '[Insert Date]', 'Review assigned subsystem findings and corrective actions'; ...
    'Product Owner', '[Insert Name]', '[Insert Date]', 'Approve vehicle-level risk and action priority'};
tbl = cell2table(rows, 'VariableNames', {'Role', 'Name', 'ReviewDate', 'Responsibility'});
end

function tbl = localTocTable(titles)
rows = cell(numel(titles), 3);
for i = 1:numel(titles)
    rows{i, 1} = char(extractBefore(titles(i), "."));
    rows{i, 2} = char(titles(i));
    rows{i, 3} = char(localSectionPurpose(titles(i)));
end
tbl = cell2table(rows, 'VariableNames', {'Section', 'Heading', 'Purpose'});
end

function tbl = localAbbreviations()
rows = {'BMS','Battery Management System'; 'RCA','Root Cause Analysis'; 'KPI','Key Performance Indicator'; ...
    'SOC','State of Charge'; 'HV','High Voltage'; 'DC','Direct Current'; 'Wh/km','Watt-hour per kilometre'; ...
    'MAE','Mean Absolute Error'; 'RMSE','Root Mean Square Error'; 'Regen','Regenerative braking or recuperation'};
tbl = cell2table(rows, 'VariableNames', {'Term', 'Meaning'});
end

function tbl = localSimulationConfig(results)
tbl = table();
if isfield(results, 'ReportOutput') && isstruct(results.ReportOutput) && isfield(results.ReportOutput, 'SimulationConfigOverview')
    tbl = localConfigToTable(results.ReportOutput.SimulationConfigOverview);
elseif isfield(results, 'SimulationConfigOverview')
    tbl = localConfigToTable(results.SimulationConfigOverview);
elseif isfield(results, 'AnalysisData') && isstruct(results.AnalysisData) && isfield(results.AnalysisData, 'Specifications') && istable(results.AnalysisData.Specifications)
    tbl = localLimit(results.AnalysisData.Specifications, 40);
end
if ~istable(tbl) || height(tbl) == 0
    tbl = cell2table({'Configuration overview','Not available','Run custom DIVe KPI/config extraction to populate this table'}, ...
        'VariableNames', {'Item', 'Value', 'EngineeringNote'});
end
end

function tbl = localConfigToTable(cfg)
rows = {};
if isstruct(cfg) && isfield(cfg, 'Sections')
    cfgSections = cfg.Sections;
    for iSection = 1:numel(cfgSections)
        titleText = localStringField(cfgSections(iSection), 'Title', "Configuration");
        if isfield(cfgSections(iSection), 'Rows')
            r = cfgSections(iSection).Rows;
            for iRow = 1:size(r, 1)
                rows(end + 1, :) = {char(titleText), char(string(r{iRow, 1})), char(string(r{iRow, 2}))}; %#ok<AGROW>
            end
        end
    end
end
if isempty(rows)
    tbl = table();
else
    tbl = cell2table(rows, 'VariableNames', {'Section', 'Parameter', 'Value'});
end
end

function tbl = localThresholdTable()
rows = {};
try
    cfg = RCA_Config();
    rows = localThresholdRows(rows, cfg, "");
catch
end
if isempty(rows)
    rows = {'Threshold configuration', 'Unavailable', 'RCA_Config could not be evaluated'};
end
tbl = cell2table(rows, 'VariableNames', {'Threshold', 'Value', 'Note'});
end

function rows = localThresholdRows(rows, value, prefix)
if isstruct(value)
    names = fieldnames(value);
    for iName = 1:numel(names)
        nextPrefix = string(names{iName});
        if strlength(prefix) > 0
            nextPrefix = prefix + "." + nextPrefix;
        end
        rows = localThresholdRows(rows, value.(names{iName}), nextPrefix);
    end
elseif isnumeric(value) || islogical(value) || ischar(value) || isstring(value)
    rows(end + 1, :) = {char(prefix), char(localTextValue(value)), 'Configurable heuristic or physical threshold'};
end
end

function tbl = localFilterKpi(kpi, tokens)
if ~istable(kpi) || height(kpi) == 0
    tbl = table();
    return;
end
tokens = lower(string(tokens));
mask = false(height(kpi), 1);
for iVar = 1:width(kpi)
    values = lower(string(kpi.(iVar)));
    for iToken = 1:numel(tokens)
        mask = mask | contains(values, tokens(iToken));
    end
end
tbl = kpi(mask, :);
if height(tbl) == 0
    tbl = kpi(1:min(height(kpi), 12), :);
end
end

function tbl = localSubsystemKpis(results)
subs = localSubsystems(results);
tbl = table();
for iSub = 1:numel(subs)
    kpi = localSubsystemKpi(subs(iSub));
    if istable(kpi) && height(kpi) > 0
        subName = repmat(cellstr(localSubsystemName(subs(iSub), iSub)), height(kpi), 1);
        kpi = addvars(kpi, subName, 'Before', 1, 'NewVariableNames', 'Subsystem');
        tbl = [tbl; kpi]; %#ok<AGROW>
    end
end
end

function tbl = localSubsystemSuggestions(results)
subs = localSubsystems(results);
rows = {};
for iSub = 1:numel(subs)
    subName = localSubsystemName(subs(iSub), iSub);
    suggestions = localStringArrayField(subs(iSub), {'Suggestions', 'RecommendationText', 'Recommendations'});
    for iSuggestion = 1:numel(suggestions)
        rows(end + 1, :) = {char(subName), char(suggestions(iSuggestion))}; %#ok<AGROW>
    end
end
if isempty(rows)
    tbl = table();
else
    tbl = cell2table(rows, 'VariableNames', {'Subsystem', 'Suggestion'});
end
end

function tbl = localFigureList(files)
files = localFigureFiles(files);
if isempty(files)
    tbl = table();
    return;
end
rows = cell(numel(files), 3);
for iFile = 1:numel(files)
    [folder, base, ext] = fileparts(char(files(iFile)));
    rows{iFile, 1} = sprintf('Figure %d', iFile);
    rows{iFile, 2} = [base ext];
    rows{iFile, 3} = folder;
end
tbl = cell2table(rows, 'VariableNames', {'FigureID', 'FileName', 'Folder'});
end

function tbl = localTableInventory(results)
rows = {};
names = {'VehicleKPI', 'SegmentSummary', 'RootCauseRanking', 'BadSegmentTable', 'OptimizationTable', ...
    'SignalPresence', 'SpecPresence', 'ExtractionLog', 'MatInventory'};
for iName = 1:numel(names)
    t = localGetTable(results, names{iName});
    if istable(t) && height(t) > 0
        rows(end + 1, :) = {names{iName}, height(t), width(t), 'Vehicle / framework result table'}; %#ok<AGROW>
    end
end
subs = localSubsystems(results);
for iSub = 1:numel(subs)
    kpi = localSubsystemKpi(subs(iSub));
    if istable(kpi) && height(kpi) > 0
        rows(end + 1, :) = {char(localSubsystemName(subs(iSub), iSub) + " KPI"), height(kpi), width(kpi), 'Subsystem result table'}; %#ok<AGROW>
    end
end
if isempty(rows)
    tbl = table();
else
    tbl = cell2table(rows, 'VariableNames', {'TableName', 'Rows', 'Columns', 'Source'});
end
end

function tbl = localScriptRefs()
folder = fileparts(mfilename('fullpath'));
files = dir(fullfile(folder, '*.m'));
rows = cell(numel(files), 3);
for iFile = 1:numel(files)
    rows{iFile, 1} = files(iFile).name;
    rows{iFile, 2} = folder;
    rows{iFile, 3} = 'RCA MATLAB source file';
end
tbl = cell2table(rows, 'VariableNames', {'Script', 'Folder', 'Purpose'});
end

function files = localAllFigures(results)
files = strings(0, 1);
if isstruct(results) && isfield(results, 'VehiclePlots') && isstruct(results.VehiclePlots) && isfield(results.VehiclePlots, 'Files')
    files = [files; localFigureFiles(results.VehiclePlots.Files)];
end
subs = localSubsystems(results);
for iSub = 1:numel(subs)
    files = [files; localSubsystemFigures(subs(iSub))]; %#ok<AGROW>
end
files = unique(localFigureFiles(files), 'stable');
end

function files = localFindFigures(results, tokens)
allFiles = localAllFigures(results);
tokens = lower(string(tokens));
if isempty(allFiles) || isempty(tokens)
    files = allFiles;
    return;
end
mask = false(numel(allFiles), 1);
for iFile = 1:numel(allFiles)
    txt = lower(string(allFiles(iFile)));
    for iToken = 1:numel(tokens)
        mask(iFile) = mask(iFile) | contains(txt, tokens(iToken));
    end
end
files = allFiles(mask);
end

function files = localSubsystemFigures(sub)
files = strings(0, 1);
fields = {'FigureFiles', 'Figures', 'PlotFiles'};
for iField = 1:numel(fields)
    if isstruct(sub) && isfield(sub, fields{iField})
        files = [files; localFigureFiles(sub.(fields{iField}))]; %#ok<AGROW>
    end
end
if isstruct(sub) && isfield(sub, 'Plots') && isstruct(sub.Plots) && isfield(sub.Plots, 'Files')
    files = [files; localFigureFiles(sub.Plots.Files)];
end
files = unique(localFigureFiles(files), 'stable');
end

function files = localFigureFiles(raw)
files = strings(0, 1);
if isempty(raw)
    return;
end
raw = string(raw(:));
for iFile = 1:numel(raw)
    p = strtrim(raw(iFile));
    if strlength(p) > 0 && isfile(p)
        files(end + 1, 1) = p; %#ok<AGROW>
    end
end
end

function subs = localSubsystems(results)
subs = struct([]);
if isstruct(results) && isfield(results, 'SubsystemResults')
    subs = results.SubsystemResults;
elseif isstruct(results) && isfield(results, 'Subsystems')
    subs = results.Subsystems;
end
if isempty(subs)
    subs = struct([]);
end
end

function kpi = localSubsystemKpi(sub)
kpi = table();
fields = {'KPITable', 'KPI', 'KpiTable', 'Kpi'};
for iField = 1:numel(fields)
    if isstruct(sub) && isfield(sub, fields{iField}) && istable(sub.(fields{iField}))
        kpi = sub.(fields{iField});
        return;
    end
end
end

function name = localSubsystemName(sub, idx)
name = "";
fields = {'Subsystem', 'SubsystemName', 'Name', 'DisplayName'};
for iField = 1:numel(fields)
    if isstruct(sub) && isfield(sub, fields{iField}) && ~isempty(sub.(fields{iField}))
        name = string(sub.(fields{iField}));
        break;
    end
end
if strlength(name) == 0
    name = "Subsystem " + string(idx);
end
name = strrep(name, "_", " ");
end

function txt = localSubsystemText(sub)
parts = strings(0, 1);
parts = localAppend(parts, localStringField(sub, 'SummaryText', ""));
parts = localAppend(parts, localStringField(sub, 'Interpretation', ""));
parts = localAppend(parts, localStringField(sub, 'Narrative', ""));
warnings = localStringArrayField(sub, {'Warnings', 'WarningText', 'Limitations'});
if ~isempty(warnings)
    parts = localAppend(parts, "Warnings / limitations:");
    for iWarn = 1:numel(warnings)
        parts = localAppend(parts, "- " + warnings(iWarn));
    end
end
suggestions = localStringArrayField(sub, {'Suggestions', 'RecommendationText', 'Recommendations'});
if ~isempty(suggestions)
    parts = localAppend(parts, "Recommended subsystem-owner actions:");
    for iSuggestion = 1:numel(suggestions)
        parts = localAppend(parts, "- " + suggestions(iSuggestion));
    end
end
if isempty(parts)
    txt = "No subsystem summary text was recorded.";
else
    txt = localJoin(parts);
end
end

function parts = localAppend(parts, value)
value = string(value);
value = value(:);
value = value(strlength(strtrim(value)) > 0);
if ~isempty(value)
    parts = [parts; value];
end
end

function value = localStringField(data, fieldName, defaultValue)
value = string(defaultValue);
if isstruct(data) && isfield(data, fieldName) && ~isempty(data.(fieldName))
    value = strjoin(localToStringArray(data.(fieldName)), newline);
end
end

function values = localStringArrayField(data, fieldNames)
values = strings(0, 1);
for iField = 1:numel(fieldNames)
    if isstruct(data) && isfield(data, fieldNames{iField}) && ~isempty(data.(fieldNames{iField}))
        values = localToStringArray(data.(fieldNames{iField}));
        values = values(strlength(strtrim(values)) > 0);
        return;
    end
end
end

function values = localToStringArray(value)
if isempty(value)
    values = strings(0, 1);
elseif isstring(value)
    values = value(:);
elseif ischar(value)
    values = string(value);
elseif istable(value)
    values = localTableRowsToText(value);
elseif iscell(value)
    values = strings(0, 1);
    for iValue = 1:numel(value)
        values = [values; localToStringArray(value{iValue})]; %#ok<AGROW>
    end
elseif isstruct(value)
    values = localStructToText(value);
elseif isnumeric(value) || islogical(value) || isdatetime(value) || isduration(value)
    values = string(value(:));
else
    try
        values = string(value(:));
    catch
        values = "[" + string(class(value)) + "]";
    end
end
values = values(:);
end

function values = localTableRowsToText(tbl)
if height(tbl) == 0
    values = strings(0, 1);
    return;
end
names = tbl.Properties.VariableNames;
values = strings(height(tbl), 1);
for iRow = 1:height(tbl)
    parts = strings(0, 1);
    for iCol = 1:width(tbl)
        cellValue = tbl{iRow, iCol};
        label = string(names{iCol});
        textValue = localTextValue(cellValue);
        if strlength(strtrim(textValue)) > 0
            parts(end + 1, 1) = label + ": " + textValue; %#ok<AGROW>
        end
    end
    if isempty(parts)
        values(iRow) = "";
    else
        values(iRow) = strjoin(parts, " | ");
    end
end
end

function values = localStructToText(data)
values = strings(0, 1);
if numel(data) ~= 1
    for iValue = 1:numel(data)
        values = [values; localStructToText(data(iValue))]; %#ok<AGROW>
    end
    return;
end
names = fieldnames(data);
for iName = 1:numel(names)
    values(end + 1, 1) = string(names{iName}) + ": " + localTextValue(data.(names{iName})); %#ok<AGROW>
end
end

function txt = localTextField(results, fieldName, fallback)
txt = string(fallback);
if isstruct(results) && isfield(results, fieldName) && ~isempty(results.(fieldName))
    txt = localJoin(localToStringArray(results.(fieldName)));
end
end

function txt = localJoin(lines)
lines = string(lines);
lines = lines(:);
if isempty(lines)
    txt = "";
else
    txt = strjoin(lines, newline);
end
end

function txt = localTextValue(value)
if isempty(value)
    txt = "";
elseif isstring(value)
    txt = strjoin(value(:), ", ");
elseif ischar(value)
    txt = string(value);
elseif isnumeric(value)
    if isscalar(value)
        if isfinite(value)
            txt = string(sprintf('%.6g', value));
        else
            txt = string(value);
        end
    else
        sz = size(value);
        txt = string(sprintf('[%d x %d %s]', sz(1), sz(2), class(value)));
    end
elseif islogical(value) || isdatetime(value)
    txt = string(value);
elseif istable(value)
    txt = strjoin(localTableRowsToText(value), "; ");
elseif iscell(value)
    pieces = strings(numel(value), 1);
    for iValue = 1:numel(value)
        pieces(iValue) = localTextValue(value{iValue});
    end
    txt = strjoin(pieces, ", ");
elseif isstruct(value)
    txt = "[struct]";
else
    try
        txt = string(value);
    catch
        txt = "[" + string(class(value)) + "]";
    end
end
end

function author = localAuthor()
author = string(getenv('USERNAME'));
if strlength(author) == 0
    author = "Simulation Engineer";
end
end

function p = localMatFile(results)
p = "";
if isstruct(results) && isfield(results, 'AnalysisData') && isstruct(results.AnalysisData) && ...
        isfield(results.AnalysisData, 'RawData') && isstruct(results.AnalysisData.RawData) && ...
        isfield(results.AnalysisData.RawData, 'MatFilePath')
    p = string(results.AnalysisData.RawData.MatFilePath);
elseif isstruct(results) && isfield(results, 'MatFilePath')
    p = string(results.MatFilePath);
elseif isstruct(results) && isfield(results, 'MatFile')
    p = string(results.MatFile);
end
if strlength(p) == 0
    p = "[Not recorded]";
end
end

function p = localExcelFile(results)
p = "";
if isstruct(results) && isfield(results, 'Metadata') && isstruct(results.Metadata) && isfield(results.Metadata, 'ExcelFile')
    p = string(results.Metadata.ExcelFile);
elseif isstruct(results) && isfield(results, 'AnalysisData') && isstruct(results.AnalysisData) && ...
        isfield(results.AnalysisData, 'Metadata') && isstruct(results.AnalysisData.Metadata) && ...
        isfield(results.AnalysisData.Metadata, 'ExcelFile')
    p = string(results.AnalysisData.Metadata.ExcelFile);
end
if strlength(p) == 0
    defaultPath = fullfile(fileparts(fileparts(mfilename('fullpath'))), 'info', 'eBus_Model_Info.xlsx');
    if isfile(defaultPath)
        p = string(defaultPath);
    else
        p = "[Not recorded]";
    end
end
end

function p = localOutputFolder(results)
p = "";
fields = {'OutputFolder', 'ResultsFolder'};
for iField = 1:numel(fields)
    if isstruct(results) && isfield(results, fields{iField}) && ~isempty(results.(fields{iField}))
        p = string(results.(fields{iField}));
        break;
    end
end
if strlength(p) == 0 && isstruct(results) && isfield(results, 'Paths') && isstruct(results.Paths)
    if isfield(results.Paths, 'OutputFolder')
        p = string(results.Paths.OutputFolder);
    elseif isfield(results.Paths, 'Output')
        p = string(results.Paths.Output);
    elseif isfield(results.Paths, 'Results')
        p = string(results.Paths.Results);
    end
end
if strlength(p) == 0
    p = pwd;
end
end

function txt = localShortSource(sourceInfo)
if strlength(string(sourceInfo.SourcePath)) > 0
    [~, n, e] = fileparts(char(sourceInfo.SourcePath));
    txt = sprintf('%s: %s%s', char(sourceInfo.SourceType), n, e);
else
    txt = char(sourceInfo.SourceType);
end
end

function name = localShortFile(filePath)
[~, n, e] = fileparts(char(string(filePath)));
name = [n e];
end

function purpose = localSectionPurpose(titleText)
t = string(titleText);
if contains(t, "Info")
    purpose = "Title, author, source data, and output context";
elseif contains(t, "Document Control")
    purpose = "Purpose, scope, version history, and review roles";
elseif contains(t, "Table of Contents")
    purpose = "Full report navigation list";
elseif contains(t, "List of Figures")
    purpose = "Available RCA figure inventory";
elseif contains(t, "List of Tables")
    purpose = "Available RCA table inventory";
elseif contains(t, "Abbreviations")
    purpose = "Nomenclature used in the RCA";
elseif contains(t, "Introduction")
    purpose = "Background, objectives, audience, and boundaries";
elseif contains(t, "Simulation")
    purpose = "Model, scenario, data source, and data-quality overview";
elseif contains(t, "Technical Summary")
    purpose = "Management-readable technical findings and actions";
elseif contains(t, "Methodology")
    purpose = "How KPIs, segments, and root causes are calculated";
elseif contains(t, "Vehicle-Level")
    purpose = "Vehicle-level performance, energy, range, balance, and operation evidence";
elseif contains(t, "Subsystem")
    purpose = "Subsystem-specific RCA and owner feedback";
elseif contains(t, "Event")
    purpose = "Acceleration, braking, cruise, grade, and worst-segment drill-down";
elseif contains(t, "KPI")
    purpose = "Dashboard tables for review";
elseif contains(t, "Root Cause")
    purpose = "Prioritized cause summary and bad-segment evidence";
elseif contains(t, "Recommendations")
    purpose = "Engineering follow-up actions";
elseif contains(t, "Conclusion")
    purpose = "Main vehicle limitations and next steps";
else
    purpose = "Supporting evidence, signal list, logs, and script references";
end
end

function localPreview(ax, filePath)
try
    cla(ax);
    img = imread(char(string(filePath)));
    image(ax, img);
    axis(ax, 'image');
    axis(ax, 'off');
    title(ax, localShortFile(filePath), 'Interpreter', 'none', 'FontWeight', 'bold');
catch exception
    cla(ax);
    text(ax, 0.5, 0.5, sprintf('Unable to preview figure:\n%s', exception.message), ...
        'HorizontalAlignment', 'center', 'Interpreter', 'none');
    axis(ax, 'off');
end
end

function localOpenFigure(filePath)
try
    img = imread(char(string(filePath)));
    f = figure('Name', localShortFile(filePath), 'Color', 'w', 'NumberTitle', 'off', 'Toolbar', 'figure');
    ax = axes('Parent', f);
    image(ax, img);
    axis(ax, 'image');
    axis(ax, 'off');
    title(ax, localShortFile(filePath), 'Interpreter', 'none');
    zoom(f, 'on');
catch exception
    warning('RCA_GUI:FigureOpenFailed', 'Unable to open figure %s: %s', char(string(filePath)), exception.message);
end
end

function localOpenFile(filePath)
try
    if ispc
        winopen(char(string(filePath)));
    else
        open(char(string(filePath)));
    end
catch exception
    warning('RCA_GUI:OpenFileFailed', 'Unable to open %s: %s', char(string(filePath)), exception.message);
end
end

function localOpenFolder(folderPath)
folderPath = char(string(folderPath));
if isempty(folderPath) || ~isfolder(folderPath)
    warning('RCA_GUI:MissingFolder', 'Folder not found: %s', folderPath);
    return;
end
try
    if ispc
        winopen(folderPath);
    else
        open(folderPath);
    end
catch exception
    warning('RCA_GUI:OpenFolderFailed', 'Unable to open %s: %s', folderPath, exception.message);
end
end

function localGenerateReport(results, language)
try
    options = struct('Language', language, 'CreateSupportingTemplateFiles', false);
    Generate_eBus_RCA_Word_Report(results, localOutputFolder(results), options);
catch exception
    try
        uialert(gcbf, exception.message, 'Word Report Generation Failed');
    catch
        warning('RCA_GUI:ReportFailed', '%s', exception.message);
    end
end
end
