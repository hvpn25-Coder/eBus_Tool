function result = Analyze_ElectricDrive(analysisData, outputPaths, config)
% Analyze_ElectricDrive  Electric machine and inverter operating analysis.

result = localInitResult("ELECTRIC DRIVE", {'emot1_pwr', 'emot2_pwr'}, ...
    {'emot1_act_trq', 'emot2_act_trq', 'emot1_act_spd', 'emot2_act_spd', 'emot1_loss_pwr', 'emot2_loss_pwr'});
t = analysisData.Derived.time_s;
elecPwr = analysisData.Derived.motorElectricalPower_kW;
mechPwr = analysisData.Derived.motorMechanicalPower_kW;
lossPwr = analysisData.Derived.motorLossPower_kW;
motorSpd = analysisData.Derived.motorSpeed_rpm;
motorTrq = analysisData.Derived.torqueActualTotal_Nm;

rows = cell(0, 7);
summary = strings(0, 1);

if all(isnan(elecPwr))
    result.Warnings(end + 1) = "Electric drive power signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Electric Drive", strings(0, 1), strings(0, 1));
    return;
end

tractiveMask = elecPwr > 1 & mechPwr > 0;
elecEnergy = trapz(t, max(elecPwr, 0)) / 3600;
mechEnergy = trapz(t, max(mechPwr, 0)) / 3600;
lossEnergy = trapz(t, max(lossPwr, 0)) / 3600;
efficiency = mechPwr ./ max(elecPwr, eps);
efficiency(~tractiveMask) = NaN;

rows = RCA_AddKPI(rows, 'Electrical Traction Energy', elecEnergy, 'kWh', 'Energy', 'Electric Drive', 'emot1_pwr + emot2_pwr', 'Integrated positive electrical power.');
rows = RCA_AddKPI(rows, 'Mechanical Traction Energy', mechEnergy, 'kWh', 'Energy', 'Electric Drive', 'torque + speed or motor power basis', 'Integrated positive mechanical output.');
rows = RCA_AddKPI(rows, 'Average Tractive Electric Drive Efficiency', mean(efficiency, 'omitnan') * 100, '%', 'Efficiency', 'Electric Drive', 'motor electrical + mechanical power', 'Only positive tractive samples are included.');
rows = RCA_AddKPI(rows, 'Electric Drive Loss Energy', lossEnergy, 'kWh', 'Losses', 'Electric Drive', 'emot1_loss_pwr + emot2_loss_pwr', 'Integrated positive loss power.');
rows = RCA_AddKPI(rows, 'Mean Motor Speed', mean(motorSpd, 'omitnan'), 'rpm', 'Operation', 'Electric Drive', 'emot1_act_spd + emot2_act_spd', 'Combined average speed.');

highSpeedShare = mean(motorSpd > config.Thresholds.HighMotorEfficiencySpeedFraction * max(motorSpd, [], 'omitnan'), 'omitnan') * 100;
rows = RCA_AddKPI(rows, 'High Motor Speed Time Share', highSpeedShare, '%', 'Efficiency', 'Electric Drive', 'motor speed', 'High-speed threshold is a heuristic fraction of observed maximum speed.');
summary(end + 1) = sprintf('Electric drive mean tractive efficiency is %.1f%% with %.2f kWh of integrated loss energy.', ...
    mean(efficiency, 'omitnan') * 100, lossEnergy);

recs = strings(0, 1);
evidence = strings(0, 1);
if mean(efficiency, 'omitnan') * 100 < 85
    recs(end + 1) = "Review motor/inverter operating region clustering; the drive spends too much tractive time away from an efficient region.";
    evidence(end + 1) = sprintf('Average tractive efficiency is %.1f%%.', mean(efficiency, 'omitnan') * 100);
end
if highSpeedShare > 15
    recs(end + 1) = "Investigate whether shift scheduling or final ratio keeps the motors in a high-speed region too often.";
    evidence(end + 1) = sprintf('Motor speed exceeds %.0f%% of observed maximum for %.1f%% of samples.', ...
        config.Thresholds.HighMotorEfficiencySpeedFraction * 100, highSpeedShare);
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, elecPwr, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, mechPwr, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, lossPwr, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Electric Drive Power Flows');
ylabel('Power (kW)');
legend({'Electrical', 'Mechanical', 'Loss'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
scatter(motorSpd, motorTrq, 12, efficiency * 100, 'filled');
title('Electric Drive Operating Region');
xlabel('Motor Speed (rpm)');
ylabel('Total Motor Torque (Nm)');
cb = colorbar;
cb.Label.String = 'Efficiency (%)';
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'ElectricDrive'), 'ElectricDrive_Overview', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Electric Drive", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
