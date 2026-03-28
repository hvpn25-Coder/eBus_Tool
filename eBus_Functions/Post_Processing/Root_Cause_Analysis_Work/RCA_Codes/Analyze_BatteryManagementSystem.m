function result = Analyze_BatteryManagementSystem(analysisData, outputPaths, config)
% Analyze_BatteryManagementSystem  Battery limit utilization analysis.

result = localInitResult("BATTERY MANAGEMENT SYSTEM", ...
    {'batt_chrg_pwr_lim', 'batt_dischrg_pwr_lim'}, {'batt_curr', 'batt_pwr', 'batt_soc', 'batt_temp'});
t = analysisData.Derived.time_s;
battPwr = analysisData.Derived.batteryPower_kW;
battCurr = analysisData.Derived.batteryCurrent_A;
chgPwrLim = analysisData.Derived.battChargePowerLimit_kW;
disPwrLim = analysisData.Derived.battDischargePowerLimit_kW;
chgCurLim = analysisData.Derived.battChargeCurrentLimit_A;
disCurLim = analysisData.Derived.battDischargeCurrentLimit_A;

rows = cell(0, 7);
summary = strings(0, 1);
if all(isnan(chgPwrLim)) && all(isnan(disPwrLim)) && all(isnan(chgCurLim)) && all(isnan(disCurLim))
    result.Warnings(end + 1) = "BMS limit signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Battery Management System", strings(0, 1), strings(0, 1));
    return;
end

dischargeLimitUse = battPwr ./ max(disPwrLim, eps);
chargeLimitUse = abs(min(battPwr, 0)) ./ max(chgPwrLim, eps);
currentLimitUse = abs(battCurr) ./ max(max(abs(chgCurLim), abs(disCurLim)), eps);

rows = RCA_AddKPI(rows, 'Discharge Power Limit Utilization Mean', mean(dischargeLimitUse, 'omitnan') * 100, '%', 'Performance', 'Battery Management System', 'batt_pwr + batt_dischrg_pwr_lim', 'Limit utilization indicates power-headroom consumption.');
rows = RCA_AddKPI(rows, 'Charge Power Limit Utilization Mean', mean(chargeLimitUse, 'omitnan') * 100, '%', 'Efficiency', 'Battery Management System', 'batt_pwr + batt_chrg_pwr_lim', 'High value can suppress regen headroom.');
rows = RCA_AddKPI(rows, 'Current Limit Utilization Mean', mean(currentLimitUse, 'omitnan') * 100, '%', 'Performance', 'Battery Management System', 'batt_curr + current limits', 'Uses the larger magnitude of charge/discharge current limits.');
rows = RCA_AddKPI(rows, 'Time Near Any Battery Limit', ...
    100 * mean(dischargeLimitUse > config.Thresholds.LimitUsageFraction | chargeLimitUse > config.Thresholds.LimitUsageFraction | currentLimitUse > config.Thresholds.LimitUsageFraction, 'omitnan'), ...
    '%', 'Performance', 'Battery Management System', 'battery power/current and limits', 'Near-limit threshold is configurable.');
summary(end + 1) = sprintf('BMS limits are active or nearly active for %.1f%% of samples.', ...
    100 * mean(dischargeLimitUse > config.Thresholds.LimitUsageFraction | chargeLimitUse > config.Thresholds.LimitUsageFraction | currentLimitUse > config.Thresholds.LimitUsageFraction, 'omitnan'));

recs = strings(0, 1);
evidence = strings(0, 1);
if mean(dischargeLimitUse > config.Thresholds.LimitUsageFraction, 'omitnan') > 0.05
    recs(end + 1) = "Discharge limits are frequently active; revisit power-limit calibration or battery capability assumptions before assigning poor acceleration only to the motor or gearbox.";
    evidence(end + 1) = "Discharge power stays near limit for more than 5% of samples.";
end
if mean(chargeLimitUse > config.Thresholds.LimitUsageFraction, 'omitnan') > 0.05
    recs(end + 1) = "Charge acceptance limits are materially restricting regen; inspect temperature and SoC dependent BMS calibration.";
    evidence(end + 1) = "Charge power stays near limit for more than 5% of samples.";
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, battPwr, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, disPwrLim, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, -chgPwrLim, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
title('Battery Power Versus Power Limits');
ylabel('Power (kW)');
legend({'Battery power', 'Discharge limit', 'Charge limit'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
plot(t, battCurr, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, disCurLim, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, -chgCurLim, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
title('Battery Current Versus Current Limits');
xlabel('Time (s)');
ylabel('Current (A)');
legend({'Battery current', 'Discharge current limit', 'Charge current limit'}, 'Location', 'best');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'BatteryManagementSystem'), 'BatteryManagementSystem_Limits', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Battery Management System", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
