function result = Analyze_Battery(analysisData, outputPaths, config)
% Analyze_Battery  Battery energy, SoC window, voltage/current behavior, and thermal RCA.

result = localInitResult("BATTERY", ...
    {'batt_pwr', 'batt_soc'}, ...
    {'batt_curr', 'batt_volt', 'batt_loss_pwr', 'batt_temp'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Battery analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Battery", recs, evidence);
    result.SummaryText = summary;
    return;
end

battPwr = d.batteryPower_kW(:);
battCurr = d.batteryCurrent_A(:);
battVolt = d.batteryVoltage_V(:);
battLoss = d.batteryLossPower_kW(:);
battSoc = d.batterySOC_pct(:);
battTemp = d.batteryTemp_C(:);
vehSpeed = d.vehVel_kmh(:);

if ~any(isfinite(battPwr))
    battPwr = localAlignedSignal(analysisData.Signals, 'batt_pwr', n);
end
if ~any(isfinite(battCurr))
    battCurr = localAlignedSignal(analysisData.Signals, 'batt_curr', n);
end
if ~any(isfinite(battVolt))
    battVolt = localAlignedSignal(analysisData.Signals, 'batt_volt', n);
end
if ~any(isfinite(battLoss))
    battLoss = localAlignedSignal(analysisData.Signals, 'batt_loss_pwr', n);
end
if ~any(isfinite(battSoc))
    battSoc = localAlignedSignal(analysisData.Signals, 'batt_soc', n);
end
if ~any(isfinite(battTemp))
    battTemp = localAlignedSignal(analysisData.Signals, 'batt_temp', n);
end

if all(isnan(battPwr)) && all(isnan(battSoc))
    result.Warnings(end + 1) = "Battery analysis skipped because battery power and SoC are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Battery", recs, evidence);
    result.SummaryText = summary;
    return;
end

movingMask = isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh;
dischargeMask = isfinite(battPwr) & battPwr > 0;
chargeMask = isfinite(battPwr) & battPwr < 0;
highDischargeMask = dischargeMask & battPwr >= 0.75 * max(battPwr, [], 'omitnan');
highChargeMask = chargeMask & abs(battPwr) >= 0.75 * max(abs(battPwr(chargeMask)), [], 'omitnan');

dischargeEnergy = RCA_TrapzFinite(t, max(battPwr, 0)) / 3600;
chargeEnergy = RCA_TrapzFinite(t, max(-battPwr, 0)) / 3600;
lossEnergy = RCA_TrapzFinite(t, max(battLoss, 0)) / 3600;

finiteSoc = battSoc(isfinite(battSoc));
if numel(finiteSoc) >= 2
    socStart = finiteSoc(1);
    socEnd = finiteSoc(end);
    socDrop = socStart - socEnd;
    socWindow = max(finiteSoc) - min(finiteSoc);
else
    socStart = NaN;
    socEnd = NaN;
    socDrop = NaN;
    socWindow = NaN;
end

voltSagPct = 100 * (max(battVolt, [], 'omitnan') - min(battVolt, [], 'omitnan')) / max(max(battVolt, [], 'omitnan'), eps);
tempRange = max(battTemp, [], 'omitnan') - min(battTemp, [], 'omitnan');
lossShare = NaN(size(battPwr));
lossShare(dischargeMask) = 100 * max(battLoss(dischargeMask), 0) ./ max(battPwr(dischargeMask), eps);

rows = RCA_AddKPI(rows, 'Battery Discharge Energy', dischargeEnergy, 'kWh', ...
    'Energy', 'Battery', 'batt_pwr', ...
    'Integrated discharge energy using RCA sign convention: discharge positive, charge negative.');
rows = RCA_AddKPI(rows, 'Battery Charge Energy', chargeEnergy, 'kWh', ...
    'Energy', 'Battery', 'batt_pwr', ...
    'Integrated charging energy magnitude using RCA sign convention.');
rows = RCA_AddKPI(rows, 'Battery Loss Energy', lossEnergy, 'kWh', ...
    'Losses', 'Battery', 'batt_loss_pwr', ...
    'Integrated positive battery loss power.');
rows = RCA_AddKPI(rows, 'Battery SoC Start', socStart, '%', ...
    'Operation', 'Battery', 'batt_soc', ...
    'Battery state of charge at trip start.');
rows = RCA_AddKPI(rows, 'Battery SoC End', socEnd, '%', ...
    'Operation', 'Battery', 'batt_soc', ...
    'Battery state of charge at trip end.');
rows = RCA_AddKPI(rows, 'Battery SoC Change', socDrop, '%', ...
    'Operation', 'Battery', 'batt_soc', ...
    'Start minus end SoC under the logged trip.');
rows = RCA_AddKPI(rows, 'Battery SoC Window Used', socWindow, '%', ...
    'Operation', 'Battery', 'batt_soc', ...
    'Maximum SoC excursion across the trip.');
rows = RCA_AddKPI(rows, 'Mean Battery Voltage', mean(battVolt, 'omitnan'), 'V', ...
    'Operation', 'Battery', 'batt_volt', ...
    'Average battery terminal voltage.');
rows = RCA_AddKPI(rows, 'Battery Voltage Sag Range', voltSagPct, '%', ...
    'Performance', 'Battery', 'batt_volt', ...
    'Voltage range normalized by peak voltage as a simple sag indicator.');
rows = RCA_AddKPI(rows, 'Peak Discharge Power', max(battPwr, [], 'omitnan'), 'kW', ...
    'Performance', 'Battery', 'batt_pwr', ...
    'Maximum discharge power using RCA sign convention.');
rows = RCA_AddKPI(rows, 'Peak Charge Power', max(-battPwr, [], 'omitnan'), 'kW', ...
    'Efficiency', 'Battery', 'batt_pwr', ...
    'Maximum charging power magnitude using RCA sign convention.');
rows = RCA_AddKPI(rows, 'Peak Discharge Current', max(battCurr, [], 'omitnan'), 'A', ...
    'Performance', 'Battery', 'batt_curr', ...
    'Maximum discharge current using RCA sign convention.');
rows = RCA_AddKPI(rows, 'Peak Charge Current', max(-battCurr, [], 'omitnan'), 'A', ...
    'Efficiency', 'Battery', 'batt_curr', ...
    'Maximum charge current magnitude using RCA sign convention.');
rows = RCA_AddKPI(rows, 'Battery Temperature Mean', mean(battTemp, 'omitnan'), 'degC', ...
    'Thermal', 'Battery', 'batt_temp', ...
    'Average battery temperature.');
rows = RCA_AddKPI(rows, 'Battery Temperature Range', tempRange, 'degC', ...
    'Thermal', 'Battery', 'batt_temp', ...
    'Battery temperature spread over the trip.');
rows = RCA_AddKPI(rows, 'Time at Low SoC', 100 * RCA_FractionTrue(battSoc <= config.Thresholds.LowSOC_pct, isfinite(battSoc)), '%', ...
    'Operation', 'Battery', 'batt_soc', ...
    sprintf('Share of samples at or below the %.1f%% low-SoC heuristic.', config.Thresholds.LowSOC_pct));
rows = RCA_AddKPI(rows, 'Time at Critical SoC', 100 * RCA_FractionTrue(battSoc <= config.Thresholds.CriticalSOC_pct, isfinite(battSoc)), '%', ...
    'Operation', 'Battery', 'batt_soc', ...
    sprintf('Share of samples at or below the %.1f%% critical-SoC heuristic.', config.Thresholds.CriticalSOC_pct));
rows = RCA_AddKPI(rows, 'High Discharge Power Share', 100 * RCA_FractionTrue(highDischargeMask, dischargeMask), '%', ...
    'Performance', 'Battery', 'batt_pwr', ...
    'Share of discharge samples in the top 25% of observed discharge power.');
rows = RCA_AddKPI(rows, 'High Charge Power Share', 100 * RCA_FractionTrue(highChargeMask, chargeMask), '%', ...
    'Efficiency', 'Battery', 'batt_pwr', ...
    'Share of charging samples in the top 25% of observed charge power magnitude.');
rows = RCA_AddKPI(rows, 'High Battery Loss Share in Discharge', 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, dischargeMask), '%', ...
    'Losses', 'Battery', 'batt_loss_pwr + batt_pwr', ...
    sprintf('Discharge samples above %.1f%% battery loss share.', config.Thresholds.HighLossShare_pct));

summary(end + 1) = sprintf(['Battery summary: %.2f kWh discharged, %.2f kWh charged, %.2f kWh lost, and %.1f%% SoC change across the trip. ', ...
    'The RCA layer uses discharge-positive and charge-negative sign convention converted from the workbook definitions.'], ...
    dischargeEnergy, chargeEnergy, lossEnergy, socDrop);
summary(end + 1) = sprintf(['Voltage and thermal context: mean voltage is %.1f V with %.1f%% sag range, and battery temperature spans %.1f degC.'], ...
    mean(battVolt, 'omitnan'), voltSagPct, tempRange);
summary(end + 1) = sprintf(['Operating window: low-SoC share is %.1f%% and critical-SoC share is %.1f%%. ', ...
    'This helps separate battery-state effects from drivetrain or control limitations.'], ...
    100 * RCA_FractionTrue(battSoc <= config.Thresholds.LowSOC_pct, isfinite(battSoc)), ...
    100 * RCA_FractionTrue(battSoc <= config.Thresholds.CriticalSOC_pct, isfinite(battSoc)));

if mean(battSoc, 'omitnan') < config.Thresholds.LowSOC_pct
    recs(end + 1) = "Low-SoC operation is material; separate low-energy-window effects from controller or drivetrain blame when diagnosing weak performance.";
    evidence(end + 1) = sprintf('Mean SoC is %.1f%%.', mean(battSoc, 'omitnan'));
end
if voltSagPct > config.Thresholds.VoltageSag_pct
    recs(end + 1) = "Investigate internal resistance, power demand peaks, and voltage-sag handling because the observed voltage spread is material.";
    evidence(end + 1) = sprintf('Voltage sag range is %.1f%% against the %.1f%% heuristic threshold.', voltSagPct, config.Thresholds.VoltageSag_pct);
end
if max(battTemp, [], 'omitnan') > 40
    recs(end + 1) = "Check battery thermal model and cooling assumptions; elevated temperature can distort resistance, capability, and ageing conclusions.";
    evidence(end + 1) = sprintf('Maximum battery temperature reached %.1f degC.', max(battTemp, [], 'omitnan'));
end
if 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, dischargeMask) > 15
    recs(end + 1) = "Battery losses are elevated for a material share of discharge operation. Review cell resistance, thermal conditions, and power demand peaks.";
    evidence(end + 1) = sprintf('High battery-loss share is %.1f%%.', 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, dischargeMask));
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'Battery');
plotFiles = localAppendPlotFile(plotFiles, localPlotOverview(figureFolder, t, battSoc, battPwr, battVolt, battCurr, battTemp, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotEnergyAndLoss(figureFolder, t, battPwr, battLoss, lossShare, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotPowerVsState(figureFolder, battSoc, battVolt, battTemp, battPwr, battCurr, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotWindowAndHistogram(figureFolder, battSoc, battVolt, battTemp, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Battery", recs, evidence);
result.FigureFiles = plotFiles;
end

function signalData = localAlignedSignal(signalStore, signalName, n)
signal = RCA_GetSignalData(signalStore, signalName);
signalData = NaN(n, 1);
if signal.Available
    if ~isempty(signal.AlignedData)
        signalData = localResizeVector(signal.AlignedData, n);
    elseif ~isempty(signal.Data)
        signalData = localResizeVector(signal.Data, n);
    end
end
end

function vector = localResizeVector(dataValue, n)
vector = NaN(n, 1);
dataValue = double(dataValue(:));
count = min(numel(dataValue), n);
if count > 0
    vector(1:count) = dataValue(1:count);
end
end

function plotFile = localPlotOverview(outputFolder, t, battSoc, battPwr, battVolt, battCurr, battTemp, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(4, 1, 1);
plot(t, battSoc, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Battery State of Charge');
ylabel('SoC (%)');
grid on;

subplot(4, 1, 2);
plot(t, battPwr, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Battery Power (Discharge +, Charge -)');
ylabel('Power (kW)');
grid on;

subplot(4, 1, 3);
plot(t, battVolt, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
yyaxis right;
plot(t, battCurr, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Battery Voltage and Current');
yyaxis left;
ylabel('Voltage (V)');
yyaxis right;
ylabel('Current (A)');
grid on;

subplot(4, 1, 4);
plot(t, battTemp, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
title('Battery Temperature');
xlabel('Time (s)');
ylabel('degC');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Battery_Overview', config));
close(fig);
end

function plotFile = localPlotEnergyAndLoss(outputFolder, t, battPwr, battLoss, lossShare, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, max(battPwr, 0), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(-battPwr, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Battery Discharge and Charge Power');
ylabel('Power (kW)');
legend({'Discharge power', 'Charge power'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, battLoss, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Battery Loss Power');
ylabel('Loss power (kW)');
grid on;

subplot(3, 1, 3);
plot(t, lossShare, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.HighLossShare_pct, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Battery Loss Share During Discharge');
xlabel('Time (s)');
ylabel('Loss share (%)');
legend({'Loss share', 'High-loss threshold'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Battery_EnergyAndLoss', config));
close(fig);
end

function plotFile = localPlotPowerVsState(outputFolder, battSoc, battVolt, battTemp, battPwr, battCurr, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid1 = isfinite(battSoc) & isfinite(battPwr);
scatter(battSoc(valid1), battPwr(valid1), 12, config.Plot.Colors.Vehicle, 'filled');
title('Battery Power Versus SoC');
xlabel('SoC (%)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 2);
valid2 = isfinite(battVolt) & isfinite(battCurr);
scatter(battVolt(valid2), battCurr(valid2), 12, config.Plot.Colors.Demand, 'filled');
title('Battery Current Versus Voltage');
xlabel('Voltage (V)');
ylabel('Current (A)');
grid on;

subplot(2, 2, 3);
valid3 = isfinite(battTemp) & isfinite(battPwr);
scatter(battTemp(valid3), battPwr(valid3), 12, config.Plot.Colors.Auxiliary, 'filled');
title('Battery Power Versus Temperature');
xlabel('Temperature (degC)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 4);
valid4 = isfinite(battSoc) & isfinite(battVolt);
scatter(battSoc(valid4), battVolt(valid4), 12, config.Plot.Colors.Battery, 'filled');
title('Battery Voltage Versus SoC');
xlabel('SoC (%)');
ylabel('Voltage (V)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Battery_PowerVsState', config));
close(fig);
end

function plotFile = localPlotWindowAndHistogram(outputFolder, battSoc, battVolt, battTemp, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(1, 3, 1);
histogram(battSoc(isfinite(battSoc)), 30, 'FaceColor', config.Plot.Colors.Battery);
title('SoC Distribution');
xlabel('SoC (%)');
ylabel('Samples');
grid on;

subplot(1, 3, 2);
histogram(battVolt(isfinite(battVolt)), 30, 'FaceColor', config.Plot.Colors.Demand);
title('Voltage Distribution');
xlabel('Voltage (V)');
ylabel('Samples');
grid on;

subplot(1, 3, 3);
histogram(battTemp(isfinite(battTemp)), 30, 'FaceColor', config.Plot.Colors.Auxiliary);
title('Temperature Distribution');
xlabel('Temperature (degC)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Battery_WindowAndHistogram', config));
close(fig);
end

function plotFiles = localAppendPlotFile(plotFiles, plotFile)
if strlength(plotFile) > 0
    plotFiles(end + 1, 1) = plotFile; %#ok<AGROW>
end
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
