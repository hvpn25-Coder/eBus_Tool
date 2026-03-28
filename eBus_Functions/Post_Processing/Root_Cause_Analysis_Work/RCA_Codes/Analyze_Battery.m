function result = Analyze_Battery(analysisData, outputPaths, config)
% Analyze_Battery  Battery energy, SoC, voltage, and thermal behaviour.

result = localInitResult("BATTERY", {'batt_pwr', 'batt_soc'}, {'batt_curr', 'batt_volt', 'batt_loss_pwr', 'batt_temp'});
t = analysisData.Derived.time_s;
battPwr = analysisData.Derived.batteryPower_kW;
soc = analysisData.Derived.batterySOC_pct;
volt = analysisData.Derived.batteryVoltage_V;
curr = analysisData.Derived.batteryCurrent_A;
loss = analysisData.Derived.batteryLossPower_kW;
temp = analysisData.Derived.batteryTemp_C;

rows = cell(0, 7);
summary = strings(0, 1);
if all(isnan(battPwr)) && all(isnan(soc))
    result.Warnings(end + 1) = "Battery power and SoC signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Battery", strings(0, 1), strings(0, 1));
    return;
end

dischargeEnergy = RCA_TrapzFinite(t, max(battPwr, 0)) / 3600;
regenEnergy = RCA_TrapzFinite(t, max(-battPwr, 0)) / 3600;
finiteSoc = soc(isfinite(soc));
if numel(finiteSoc) >= 2
    socDrop = max(finiteSoc(1) - finiteSoc(end), 0);
else
    socDrop = NaN;
end
voltSagPct = 100 * (max(volt, [], 'omitnan') - min(volt, [], 'omitnan')) / max(max(volt, [], 'omitnan'), eps);

rows = RCA_AddKPI(rows, 'Battery Discharge Energy', dischargeEnergy, 'kWh', 'Energy', 'Battery', 'batt_pwr', 'Integrated discharge-positive battery power after workbook sign normalization.');
rows = RCA_AddKPI(rows, 'Battery Regen Energy', regenEnergy, 'kWh', 'Energy', 'Battery', 'batt_pwr', 'Integrated charging/recovered battery power after workbook sign normalization.');
rows = RCA_AddKPI(rows, 'Battery SoC Drop', socDrop, '%', 'Operation', 'Battery', 'batt_soc', 'Trip SoC consumption.');
rows = RCA_AddKPI(rows, 'Mean Battery Voltage', mean(volt, 'omitnan'), 'V', 'Operation', 'Battery', 'batt_volt', 'Average terminal voltage.');
rows = RCA_AddKPI(rows, 'Battery Voltage Sag Range', voltSagPct, '%', 'Performance', 'Battery', 'batt_volt', 'Voltage span across the trip as a simple sag indicator.');
rows = RCA_AddKPI(rows, 'Mean Battery Current', mean(curr, 'omitnan'), 'A', 'Operation', 'Battery', 'batt_curr', 'Average battery current in RCA convention: discharge positive, charge negative.');
rows = RCA_AddKPI(rows, 'Battery Loss Energy', RCA_TrapzFinite(t, max(loss, 0)) / 3600, 'kWh', 'Losses', 'Battery', 'batt_loss_pwr', 'Integrated positive battery loss power.');
rows = RCA_AddKPI(rows, 'Battery Temperature Range', max(temp, [], 'omitnan') - min(temp, [], 'omitnan'), 'degC', 'Operation', 'Battery', 'batt_temp', 'Temperature spread across the trip.');
summary(end + 1) = sprintf('Battery summary: %.2f kWh discharged, %.2f kWh recovered, %.1f%% SoC drop. RCA sign convention uses discharge positive and charge negative based on workbook metadata.', ...
    dischargeEnergy, regenEnergy, socDrop);

recs = strings(0, 1);
evidence = strings(0, 1);
if mean(soc, 'omitnan') < config.Thresholds.LowSOC_pct
    recs(end + 1) = "Low-SoC operation is present; separate low-energy-window effects from controller or gearbox blame in poor-performance segments.";
    evidence(end + 1) = sprintf('Mean SoC is %.1f%%.', mean(soc, 'omitnan'));
end
if voltSagPct > config.Thresholds.VoltageSag_pct
    recs(end + 1) = "Investigate battery internal resistance, current demand peaks, and voltage-sag handling; the observed voltage spread is material.";
    evidence(end + 1) = sprintf('Voltage sag range is %.1f%% against a %.1f%% heuristic threshold.', voltSagPct, config.Thresholds.VoltageSag_pct);
end
if max(temp, [], 'omitnan') > 40
    recs(end + 1) = "Check battery thermal model and cooling demand; elevated temperature can bias resistance, limits, and ageing conclusions.";
    evidence(end + 1) = sprintf('Maximum battery temperature reached %.1f degC.', max(temp, [], 'omitnan'));
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(4, 1, 1);
plot(t, soc, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Battery State of Charge');
ylabel('SoC (%)');
grid on;

subplot(4, 1, 2);
plot(t, battPwr, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Battery Power (Discharge +, Charge -)');
ylabel('Power (kW)');
grid on;

subplot(4, 1, 3);
plot(t, volt, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, curr, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Battery Voltage and Current (Discharge Current +)');
ylabel('V / A');
legend({'Voltage', 'Current'}, 'Location', 'best');
grid on;

subplot(4, 1, 4);
plot(t, temp, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
title('Battery Temperature');
xlabel('Time (s)');
ylabel('degC');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Battery'), 'Battery_Overview', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Battery", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
