function result = Analyze_BatteryManagementSystem(analysisData, outputPaths, config)
% Analyze_BatteryManagementSystem  Charge/discharge limit behavior and battery capability RCA.

result = localInitResult("BATTERY MANAGEMENT SYSTEM", ...
    {'batt_chrg_pwr_lim', 'batt_dischrg_pwr_lim'}, ...
    {'batt_chrg_curr_lim', 'batt_dischrg_curr_lim', 'batt_pwr', 'batt_curr', 'batt_soc', 'batt_temp'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Battery management system analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Battery Management System", recs, evidence);
    result.SummaryText = summary;
    return;
end

battPwr = d.batteryPower_kW(:);
battCurr = d.batteryCurrent_A(:);
battSoc = d.batterySOC_pct(:);
battTemp = d.batteryTemp_C(:);
chgPwrLim = d.battChargePowerLimit_kW(:);
disPwrLim = d.battDischargePowerLimit_kW(:);
chgCurLim = d.battChargeCurrentLimit_A(:);
disCurLim = d.battDischargeCurrentLimit_A(:);

if ~any(isfinite(chgPwrLim))
    chgPwrLim = localAlignedSignal(analysisData.Signals, 'batt_chrg_pwr_lim', n);
end
if ~any(isfinite(disPwrLim))
    disPwrLim = localAlignedSignal(analysisData.Signals, 'batt_dischrg_pwr_lim', n);
end
if ~any(isfinite(chgCurLim))
    chgCurLim = localAlignedSignal(analysisData.Signals, 'batt_chrg_curr_lim', n);
end
if ~any(isfinite(disCurLim))
    disCurLim = localAlignedSignal(analysisData.Signals, 'batt_dischrg_curr_lim', n);
end

if all(isnan(chgPwrLim)) && all(isnan(disPwrLim)) && all(isnan(chgCurLim)) && all(isnan(disCurLim))
    result.Warnings(end + 1) = "BMS analysis skipped because charge/discharge limit signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Battery Management System", recs, evidence);
    result.SummaryText = summary;
    return;
end

dischargePowerActive = isfinite(battPwr) & isfinite(disPwrLim) & battPwr > 0 & disPwrLim > 0;
chargePowerActive = isfinite(battPwr) & isfinite(chgPwrLim) & battPwr < 0 & chgPwrLim > 0;
dischargeCurrentActive = isfinite(battCurr) & isfinite(disCurLim) & battCurr > 0 & disCurLim > 0;
chargeCurrentActive = isfinite(battCurr) & isfinite(chgCurLim) & battCurr < 0 & chgCurLim > 0;

dischargePowerUse = NaN(size(battPwr));
chargePowerUse = NaN(size(battPwr));
dischargeCurrentUse = NaN(size(battCurr));
chargeCurrentUse = NaN(size(battCurr));
dischargePowerReserve = NaN(size(battPwr));
chargePowerReserve = NaN(size(battPwr));
dischargeCurrentReserve = NaN(size(battCurr));
chargeCurrentReserve = NaN(size(battCurr));

dischargePowerUse(dischargePowerActive) = battPwr(dischargePowerActive) ./ disPwrLim(dischargePowerActive);
chargePowerUse(chargePowerActive) = -battPwr(chargePowerActive) ./ chgPwrLim(chargePowerActive);
dischargeCurrentUse(dischargeCurrentActive) = battCurr(dischargeCurrentActive) ./ disCurLim(dischargeCurrentActive);
chargeCurrentUse(chargeCurrentActive) = -battCurr(chargeCurrentActive) ./ chgCurLim(chargeCurrentActive);

dischargePowerReserve(dischargePowerActive) = disPwrLim(dischargePowerActive) - battPwr(dischargePowerActive);
chargePowerReserve(chargePowerActive) = chgPwrLim(chargePowerActive) - abs(battPwr(chargePowerActive));
dischargeCurrentReserve(dischargeCurrentActive) = disCurLim(dischargeCurrentActive) - battCurr(dischargeCurrentActive);
chargeCurrentReserve(chargeCurrentActive) = chgCurLim(chargeCurrentActive) - abs(battCurr(chargeCurrentActive));

nearDischargePower = dischargePowerActive & dischargePowerUse >= config.Thresholds.LimitUsageFraction;
nearChargePower = chargePowerActive & chargePowerUse >= config.Thresholds.LimitUsageFraction;
nearDischargeCurrent = dischargeCurrentActive & dischargeCurrentUse >= config.Thresholds.LimitUsageFraction;
nearChargeCurrent = chargeCurrentActive & chargeCurrentUse >= config.Thresholds.LimitUsageFraction;
validAnyLimit = dischargePowerActive | chargePowerActive | dischargeCurrentActive | chargeCurrentActive;
nearAnyLimit = nearDischargePower | nearChargePower | nearDischargeCurrent | nearChargeCurrent;

rows = RCA_AddKPI(rows, 'Discharge Power Limit Utilization Mean', mean(dischargePowerUse, 'omitnan') * 100, '%', ...
    'Capability', 'Battery Management System', 'batt_pwr + batt_dischrg_pwr_lim', ...
    'Mean utilization over discharge-active samples only. RCA uses discharge-positive power internally.');
rows = RCA_AddKPI(rows, 'Charge Power Limit Utilization Mean', mean(chargePowerUse, 'omitnan') * 100, '%', ...
    'Capability', 'Battery Management System', 'batt_pwr + batt_chrg_pwr_lim', ...
    'Mean utilization over charge-active samples only. RCA uses charge-negative power internally and evaluates magnitude.');
rows = RCA_AddKPI(rows, 'Discharge Current Limit Utilization Mean', mean(dischargeCurrentUse, 'omitnan') * 100, '%', ...
    'Capability', 'Battery Management System', 'batt_curr + batt_dischrg_curr_lim', ...
    'Mean discharge-current limit utilization over discharge-active samples.');
rows = RCA_AddKPI(rows, 'Charge Current Limit Utilization Mean', mean(chargeCurrentUse, 'omitnan') * 100, '%', ...
    'Capability', 'Battery Management System', 'batt_curr + batt_chrg_curr_lim', ...
    'Mean charge-current limit utilization over charge-active samples.');
rows = RCA_AddKPI(rows, 'Time Near Any Battery Limit', 100 * RCA_FractionTrue(nearAnyLimit, validAnyLimit), '%', ...
    'Capability', 'Battery Management System', 'battery power/current and BMS limits', ...
    'Near-limit threshold is configurable and evaluated only on active valid samples.');
rows = RCA_AddKPI(rows, 'Near Discharge Power Limit Share', 100 * RCA_FractionTrue(nearDischargePower, dischargePowerActive), '%', ...
    'Capability', 'Battery Management System', 'batt_pwr + batt_dischrg_pwr_lim', ...
    'Share of discharge-active samples near discharge power limit.');
rows = RCA_AddKPI(rows, 'Near Charge Power Limit Share', 100 * RCA_FractionTrue(nearChargePower, chargePowerActive), '%', ...
    'Capability', 'Battery Management System', 'batt_pwr + batt_chrg_pwr_lim', ...
    'Share of charge-active samples near charge power limit.');
rows = RCA_AddKPI(rows, 'Near Discharge Current Limit Share', 100 * RCA_FractionTrue(nearDischargeCurrent, dischargeCurrentActive), '%', ...
    'Capability', 'Battery Management System', 'batt_curr + batt_dischrg_curr_lim', ...
    'Share of discharge-current-active samples near discharge current limit.');
rows = RCA_AddKPI(rows, 'Near Charge Current Limit Share', 100 * RCA_FractionTrue(nearChargeCurrent, chargeCurrentActive), '%', ...
    'Capability', 'Battery Management System', 'batt_curr + batt_chrg_curr_lim', ...
    'Share of charge-current-active samples near charge current limit.');
rows = RCA_AddKPI(rows, 'Mean Discharge Power Reserve', mean(dischargePowerReserve, 'omitnan'), 'kW', ...
    'Capability', 'Battery Management System', 'batt_dischrg_pwr_lim - batt_pwr', ...
    'Average power headroom during discharge-active samples.');
rows = RCA_AddKPI(rows, 'Mean Charge Power Reserve', mean(chargePowerReserve, 'omitnan'), 'kW', ...
    'Capability', 'Battery Management System', 'batt_chrg_pwr_lim - abs(batt_pwr)', ...
    'Average charge-acceptance headroom during charge-active samples.');
rows = RCA_AddKPI(rows, 'Mean Discharge Current Reserve', mean(dischargeCurrentReserve, 'omitnan'), 'A', ...
    'Capability', 'Battery Management System', 'batt_dischrg_curr_lim - batt_curr', ...
    'Average current headroom during discharge-active samples.');
rows = RCA_AddKPI(rows, 'Mean Charge Current Reserve', mean(chargeCurrentReserve, 'omitnan'), 'A', ...
    'Capability', 'Battery Management System', 'batt_chrg_curr_lim - abs(batt_curr)', ...
    'Average current headroom during charge-active samples.');
rows = RCA_AddKPI(rows, 'Low-SoC Near-Limit Share', 100 * RCA_FractionTrue(nearAnyLimit, validAnyLimit & battSoc <= config.Thresholds.LowSOC_pct), '%', ...
    'Context', 'Battery Management System', 'battery limits + batt_soc', ...
    'Share of low-SoC active-limit samples that are near any limit.');
rows = RCA_AddKPI(rows, 'Hot-Battery Near-Limit Share', 100 * RCA_FractionTrue(nearAnyLimit, validAnyLimit & battTemp >= 40), '%', ...
    'Context', 'Battery Management System', 'battery limits + batt_temp', ...
    'Share of hot-battery active-limit samples that are near any limit.');

summary(end + 1) = sprintf(['BMS limits are near-active for %.1f%% of active limit samples. ', ...
    'The RCA layer interprets charge/discharge raw signs from the workbook and converts them to discharge-positive internal convention.'], ...
    100 * RCA_FractionTrue(nearAnyLimit, validAnyLimit));
summary(end + 1) = sprintf(['Mean discharge power utilization is %.1f%%, charge power utilization is %.1f%%, and current-limit utilization ranges from %.1f%% to %.1f%%.'], ...
    mean(dischargePowerUse, 'omitnan') * 100, mean(chargePowerUse, 'omitnan') * 100, ...
    mean(dischargeCurrentUse, 'omitnan') * 100, mean(chargeCurrentUse, 'omitnan') * 100);
summary(end + 1) = sprintf(['Limit-context summary: low-SoC near-limit share is %.1f%% and hot-battery near-limit share is %.1f%%. ', ...
    'This helps identify whether capability loss aligns with battery state or temperature.'], ...
    100 * RCA_FractionTrue(nearAnyLimit, validAnyLimit & battSoc <= config.Thresholds.LowSOC_pct), ...
    100 * RCA_FractionTrue(nearAnyLimit, validAnyLimit & battTemp >= 40));

if RCA_FractionTrue(nearDischargePower, dischargePowerActive) > 0.05
    recs(end + 1) = "Discharge power limits are frequently active. Review power-limit calibration or battery capability assumptions before assigning weak acceleration only to motor or gearbox.";
    evidence(end + 1) = sprintf('Near discharge power limit share is %.1f%%.', 100 * RCA_FractionTrue(nearDischargePower, dischargePowerActive));
end
if RCA_FractionTrue(nearChargePower, chargePowerActive) > 0.05
    recs(end + 1) = "Charge acceptance limits are materially restricting regen. Inspect temperature and SoC dependent BMS calibration.";
    evidence(end + 1) = sprintf('Near charge power limit share is %.1f%%.', 100 * RCA_FractionTrue(nearChargePower, chargePowerActive));
end
if RCA_FractionTrue(nearDischargeCurrent, dischargeCurrentActive) > 0.05 || RCA_FractionTrue(nearChargeCurrent, chargeCurrentActive) > 0.05
    recs(end + 1) = "Current limits are frequently active. Review current-limit shaping and whether voltage or thermal protection is constraining capability earlier than expected.";
    evidence(end + 1) = sprintf('Near discharge current limit share is %.1f%% and near charge current limit share is %.1f%%.', ...
        100 * RCA_FractionTrue(nearDischargeCurrent, dischargeCurrentActive), 100 * RCA_FractionTrue(nearChargeCurrent, chargeCurrentActive));
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'BatteryManagementSystem');
plotFiles = localAppendPlotFile(plotFiles, localPlotPowerLimits(figureFolder, t, battPwr, disPwrLim, chgPwrLim, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotCurrentLimits(figureFolder, t, battCurr, disCurLim, chgCurLim, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotLimitUtilization(figureFolder, t, dischargePowerUse, chargePowerUse, dischargeCurrentUse, chargeCurrentUse, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotLimitContext(figureFolder, battSoc, battTemp, battPwr, battCurr, nearAnyLimit, validAnyLimit, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Battery Management System", recs, evidence);
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

function plotFile = localPlotPowerLimits(outputFolder, t, battPwr, disPwrLim, chgPwrLim, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

plot(t, battPwr, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, disPwrLim, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, -chgPwrLim, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Battery Power Versus Power Limits (Discharge +, Charge -)');
xlabel('Time (s)');
ylabel('Power (kW)');
legend({'Battery power', 'Discharge power limit', 'Charge power limit'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'BatteryManagementSystem_PowerLimits', config));
close(fig);
end

function plotFile = localPlotCurrentLimits(outputFolder, t, battCurr, disCurLim, chgCurLim, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

plot(t, battCurr, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, disCurLim, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, -chgCurLim, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Battery Current Versus Current Limits (Discharge +, Charge -)');
xlabel('Time (s)');
ylabel('Current (A)');
legend({'Battery current', 'Discharge current limit', 'Charge current limit'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'BatteryManagementSystem_CurrentLimits', config));
close(fig);
end

function plotFile = localPlotLimitUtilization(outputFolder, t, dischargePowerUse, chargePowerUse, dischargeCurrentUse, chargeCurrentUse, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 1, 1);
plot(t, dischargePowerUse * 100, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, chargePowerUse * 100, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.LimitUsageFraction * 100, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Battery Power Limit Utilization');
ylabel('Utilization (%)');
legend({'Discharge power use', 'Charge power use', 'Near-limit threshold'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
plot(t, dischargeCurrentUse * 100, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, chargeCurrentUse * 100, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.LimitUsageFraction * 100, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Battery Current Limit Utilization');
xlabel('Time (s)');
ylabel('Utilization (%)');
legend({'Discharge current use', 'Charge current use', 'Near-limit threshold'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'BatteryManagementSystem_LimitUtilization', config));
close(fig);
end

function plotFile = localPlotLimitContext(outputFolder, battSoc, battTemp, battPwr, battCurr, nearAnyLimit, validAnyLimit, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid1 = isfinite(battSoc) & isfinite(battPwr);
scatter(battSoc(valid1), battPwr(valid1), 12, double(nearAnyLimit(valid1 & validAnyLimit)), 'filled');
colormap(gca, [config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
title('Battery Power Versus SoC');
xlabel('SoC (%)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 2);
valid2 = isfinite(battTemp) & isfinite(battPwr);
scatter(battTemp(valid2), battPwr(valid2), 12, double(nearAnyLimit(valid2 & validAnyLimit)), 'filled');
colormap(gca, [config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
title('Battery Power Versus Temperature');
xlabel('Temperature (degC)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 3);
valid3 = isfinite(battSoc) & isfinite(battCurr);
scatter(battSoc(valid3), battCurr(valid3), 12, double(nearAnyLimit(valid3 & validAnyLimit)), 'filled');
colormap(gca, [config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
title('Battery Current Versus SoC');
xlabel('SoC (%)');
ylabel('Current (A)');
grid on;

subplot(2, 2, 4);
histogram(double(nearAnyLimit(validAnyLimit)), 'FaceColor', config.Plot.Colors.Warning);
title('Near-Limit Event Histogram');
xlabel('0 = Not near limit, 1 = Near limit');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'BatteryManagementSystem_LimitContext', config));
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
