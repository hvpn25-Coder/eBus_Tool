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

motorSpdAbs = abs(motorSpd);
tractiveMask = isfinite(elecPwr) & isfinite(mechPwr) & elecPwr > 1 & mechPwr > 0;
elecEnergy = RCA_TrapzFinite(t, max(elecPwr, 0)) / 3600;
mechEnergy = RCA_TrapzFinite(t, max(mechPwr, 0)) / 3600;
lossEnergy = RCA_TrapzFinite(t, max(lossPwr, 0)) / 3600;
efficiency = NaN(size(elecPwr));
efficiency(tractiveMask) = mechPwr(tractiveMask) ./ max(elecPwr(tractiveMask), eps);
avgEffPct = mean(efficiency, 'omitnan') * 100;

rows = RCA_AddKPI(rows, 'Electrical Traction Energy', elecEnergy, 'kWh', 'Energy', 'Electric Drive', 'emot1_pwr + emot2_pwr', 'Integrated drive-positive electrical power after workbook sign normalization.');
rows = RCA_AddKPI(rows, 'Mechanical Traction Energy', mechEnergy, 'kWh', 'Energy', 'Electric Drive', 'torque + speed or motor power basis', 'Integrated positive mechanical output.');
rows = RCA_AddKPI(rows, 'Average Tractive Electric Drive Efficiency', avgEffPct, '%', 'Efficiency', 'Electric Drive', 'motor electrical + mechanical power', 'Only positive tractive samples are included.');
rows = RCA_AddKPI(rows, 'Electric Drive Loss Energy', lossEnergy, 'kWh', 'Losses', 'Electric Drive', 'emot1_loss_pwr + emot2_loss_pwr', 'Integrated positive loss power.');
rows = RCA_AddKPI(rows, 'Mean Motor Speed Magnitude', mean(motorSpdAbs, 'omitnan'), 'rpm', 'Operation', 'Electric Drive', 'emot1_act_spd + emot2_act_spd', 'Combined average speed magnitude.');

validHighSpeed = tractiveMask & isfinite(motorSpdAbs);
if any(validHighSpeed)
    highSpeedRef = max(motorSpdAbs(validHighSpeed), [], 'omitnan');
    highSpeedShare = 100 * RCA_FractionTrue(motorSpdAbs > config.Thresholds.HighMotorEfficiencySpeedFraction * highSpeedRef, validHighSpeed);
else
    highSpeedShare = NaN;
end
rows = RCA_AddKPI(rows, 'High Motor Speed Time Share', highSpeedShare, '%', 'Efficiency', 'Electric Drive', 'motor speed', 'High-speed threshold is a heuristic fraction of observed maximum speed.');
summary(end + 1) = sprintf('Electric drive mean tractive efficiency is %.1f%% with %.2f kWh of integrated loss energy.', ...
    avgEffPct, lossEnergy);

recs = strings(0, 1);
evidence = strings(0, 1);
if avgEffPct < 85
    recs(end + 1) = "Review motor/inverter operating region clustering; the drive spends too much tractive time away from an efficient region.";
    evidence(end + 1) = sprintf('Average tractive efficiency is %.1f%%.', avgEffPct);
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
validScatter = isfinite(motorSpdAbs) & isfinite(motorTrq) & isfinite(efficiency);
scatter(motorSpdAbs(validScatter), motorTrq(validScatter), 12, efficiency(validScatter) * 100, 'filled');
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
