function result = Analyze_PneumaticBrakeSystem(analysisData, outputPaths, config)
% Analyze_PneumaticBrakeSystem  Friction braking and regen opportunity analysis.

result = localInitResult("PNEUMATIC BRAKE SYSTEM", {'fric_brk_pwr'}, {'fric_brk_force', 'brk_pdl', 'batt_pwr'});
t = analysisData.Derived.time_s;
fricPwr = analysisData.Derived.frictionBrakePower_kW;
battPwr = analysisData.Derived.batteryPower_kW;
vehSpeed = analysisData.Derived.vehVel_kmh;

rows = cell(0, 7);
summary = strings(0, 1);
if all(isnan(fricPwr))
    result.Warnings(end + 1) = "Friction brake power signal is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Pneumatic Brake System", strings(0, 1), strings(0, 1));
    return;
end

fricEnergy = RCA_TrapzFinite(t, max(fricPwr, 0)) / 3600;
regenEnergy = RCA_TrapzFinite(t, max(-battPwr, 0)) / 3600;
brakingOpportunity = fricEnergy + regenEnergy;
if brakingOpportunity > 0
    regenRecovery = 100 * regenEnergy / brakingOpportunity;
else
    regenRecovery = NaN;
end
validBrakeOpportunity = isfinite(fricPwr) & isfinite(vehSpeed);
rows = RCA_AddKPI(rows, 'Friction Brake Energy', fricEnergy, 'kWh', 'Losses', 'Pneumatic Brake System', 'fric_brk_pwr', 'Integrated positive friction brake power.');
rows = RCA_AddKPI(rows, 'Approximate Regen Recovery Fraction', regenRecovery, '%', 'Efficiency', 'Pneumatic Brake System', 'fric_brk_pwr + batt_pwr', 'Approximated from electrical recovery and friction dissipation after workbook sign normalization.');
rows = RCA_AddKPI(rows, 'Brake Power Above Regen Opportunity Threshold', ...
    100 * RCA_FractionTrue(fricPwr > config.Thresholds.RegenOpportunityBrakePower_kW & vehSpeed > config.Thresholds.StopSpeed_kmh, validBrakeOpportunity), ...
    '%', 'Efficiency', 'Pneumatic Brake System', 'fric_brk_pwr + veh_vel', 'Heuristic threshold from RCA_Config.');
summary(end + 1) = sprintf('Brake system dissipates %.2f kWh as friction heat and recovers about %.1f%% of the observable braking opportunity. Recovered battery power uses workbook charge/discharge sign convention.', ...
    fricEnergy, regenRecovery);

recs = strings(0, 1);
evidence = strings(0, 1);
if regenRecovery < config.Thresholds.PoorRegenRecoveryFraction * 100
    recs(end + 1) = "Improve brake blending or regen availability; friction braking is consuming a large share of recoverable braking energy.";
    evidence(end + 1) = sprintf('Approximate regen recovery is %.1f%%.', regenRecovery);
end
if fricEnergy > 0.5
    recs(end + 1) = "Inspect deceleration calibration and battery charge acceptance limits during braking-heavy segments.";
    evidence(end + 1) = sprintf('Integrated friction brake energy is %.2f kWh.', fricEnergy);
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, fricPwr, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Friction Brake Power');
ylabel('Power (kW)');
grid on;

subplot(2, 1, 2);
plot(t, -min(battPwr, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(fricPwr, 0), '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Recovered Electrical Braking Versus Friction Braking');
xlabel('Time (s)');
ylabel('Power (kW)');
legend({'Recovered electrical power', 'Friction brake power'}, 'Location', 'best');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'PneumaticBrakeSystem'), 'PneumaticBrakeSystem_Regen', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Pneumatic Brake System", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
