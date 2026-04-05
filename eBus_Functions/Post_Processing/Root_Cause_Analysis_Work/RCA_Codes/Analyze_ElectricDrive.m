function result = Analyze_ElectricDrive(analysisData, outputPaths, config)
% Analyze_ElectricDrive  Electric motor and inverter RCA using per-machine operating signals.

result = localInitResult("ELECTRIC DRIVE", ...
    {'emot1_act_trq', 'emot2_act_trq', 'emot1_act_spd', 'emot2_act_spd', 'emot1_pwr', 'emot2_pwr'}, ...
    {'emot1_loss_pwr', 'emot2_loss_pwr', 'emot1_max_av_trq', 'emot2_max_av_trq', 'emot1_min_av_trq', 'emot2_min_av_trq'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Electric drive analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Electric Drive", recs, evidence);
    result.SummaryText = summary;
    return;
end

em1Trq = localAlignedSignal(analysisData.Signals, 'emot1_act_trq', n);
em2Trq = localAlignedSignal(analysisData.Signals, 'emot2_act_trq', n);
em1SpdRad = localAlignedSignal(analysisData.Signals, 'emot1_act_spd', n);
em2SpdRad = localAlignedSignal(analysisData.Signals, 'emot2_act_spd', n);
em1ElecPwr = localAlignedSignal(analysisData.Signals, 'emot1_pwr', n);
em2ElecPwr = localAlignedSignal(analysisData.Signals, 'emot2_pwr', n);
em1LossPwr = localAlignedSignal(analysisData.Signals, 'emot1_loss_pwr', n);
em2LossPwr = localAlignedSignal(analysisData.Signals, 'emot2_loss_pwr', n);
em1Max = localAlignedSignal(analysisData.Signals, 'emot1_max_av_trq', n);
em2Max = localAlignedSignal(analysisData.Signals, 'emot2_max_av_trq', n);
em1Min = localAlignedSignal(analysisData.Signals, 'emot1_min_av_trq', n);
em2Min = localAlignedSignal(analysisData.Signals, 'emot2_min_av_trq', n);

em1MechPwr = em1Trq .* em1SpdRad / 1000;
em2MechPwr = em2Trq .* em2SpdRad / 1000;
totalElecPwr = em1ElecPwr + em2ElecPwr;
totalMechPwr = em1MechPwr + em2MechPwr;
totalLossPwr = em1LossPwr + em2LossPwr;

em1SpdRpm = abs(em1SpdRad) * 60 / (2 * pi);
em2SpdRpm = abs(em2SpdRad) * 60 / (2 * pi);
vehSpeed = d.vehVel_kmh(:);

if ~any(isfinite(totalElecPwr))
    result.Warnings(end + 1) = "Electric drive analysis skipped because motor electrical power signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Electric Drive", recs, evidence);
    result.SummaryText = summary;
    return;
end

activeThreshold = 1.0;
driveMask = isfinite(totalElecPwr) & isfinite(totalMechPwr) & totalElecPwr > activeThreshold & totalMechPwr > 0;
regenMask = isfinite(totalElecPwr) & isfinite(totalMechPwr) & totalElecPwr < -activeThreshold & totalMechPwr < 0;

em1DriveMask = isfinite(em1ElecPwr) & isfinite(em1MechPwr) & em1ElecPwr > activeThreshold & em1MechPwr > 0;
em2DriveMask = isfinite(em2ElecPwr) & isfinite(em2MechPwr) & em2ElecPwr > activeThreshold & em2MechPwr > 0;
em1RegenMask = isfinite(em1ElecPwr) & isfinite(em1MechPwr) & em1ElecPwr < -activeThreshold & em1MechPwr < 0;
em2RegenMask = isfinite(em2ElecPwr) & isfinite(em2MechPwr) & em2ElecPwr < -activeThreshold & em2MechPwr < 0;

tractiveEfficiency = NaN(size(totalElecPwr));
tractiveEfficiency(driveMask) = totalMechPwr(driveMask) ./ max(totalElecPwr(driveMask), eps);
regenEfficiency = NaN(size(totalElecPwr));
regenEfficiency(regenMask) = abs(totalElecPwr(regenMask)) ./ max(abs(totalMechPwr(regenMask)), eps);

em1EfficiencyDrive = localEfficiency(em1MechPwr, em1ElecPwr, em1DriveMask);
em2EfficiencyDrive = localEfficiency(em2MechPwr, em2ElecPwr, em2DriveMask);
em1EfficiencyRegen = localRegenEfficiency(em1MechPwr, em1ElecPwr, em1RegenMask);
em2EfficiencyRegen = localRegenEfficiency(em2MechPwr, em2ElecPwr, em2RegenMask);

lossShare = NaN(size(totalElecPwr));
lossShare(driveMask) = 100 * totalLossPwr(driveMask) ./ max(totalElecPwr(driveMask), eps);

validPosLimit = driveMask & isfinite(d.torquePositiveLimit_Nm) & d.torquePositiveLimit_Nm > 0;
validNegLimit = regenMask & isfinite(d.torqueNegativeLimit_Nm) & abs(d.torqueNegativeLimit_Nm) > 0;
driveLimitUsage = NaN(size(totalElecPwr));
regenLimitUsage = NaN(size(totalElecPwr));
driveLimitUsage(validPosLimit) = 100 * d.torqueActualTotal_Nm(validPosLimit) ./ d.torquePositiveLimit_Nm(validPosLimit);
regenLimitUsage(validNegLimit) = 100 * abs(d.torqueActualTotal_Nm(validNegLimit)) ./ abs(d.torqueNegativeLimit_Nm(validNegLimit));
nearDriveLimitMask = validPosLimit & driveLimitUsage >= config.Thresholds.LimitUsageFraction * 100;
nearRegenLimitMask = validNegLimit & regenLimitUsage >= config.Thresholds.LimitUsageFraction * 100;

em1NearDrive = localMotorNearLimit(em1Trq, em1Max, em1DriveMask, config.Thresholds.LimitUsageFraction, true);
em2NearDrive = localMotorNearLimit(em2Trq, em2Max, em2DriveMask, config.Thresholds.LimitUsageFraction, true);
em1NearRegen = localMotorNearLimit(em1Trq, em1Min, em1RegenMask, config.Thresholds.LimitUsageFraction, false);
em2NearRegen = localMotorNearLimit(em2Trq, em2Min, em2RegenMask, config.Thresholds.LimitUsageFraction, false);

em1TorqueReserveDrive = em1Max - em1Trq;
em2TorqueReserveDrive = em2Max - em2Trq;
em1TorqueReserveRegen = abs(em1Min) - abs(em1Trq);
em2TorqueReserveRegen = abs(em2Min) - abs(em2Trq);

highSpeedRef = max([em1SpdRpm(:); em2SpdRpm(:)], [], 'omitnan');
em1HighSpeedMask = isfinite(em1SpdRpm) & em1SpdRpm >= config.Thresholds.HighMotorEfficiencySpeedFraction * highSpeedRef;
em2HighSpeedMask = isfinite(em2SpdRpm) & em2SpdRpm >= config.Thresholds.HighMotorEfficiencySpeedFraction * highSpeedRef;
lowSpeedMask = isfinite(vehSpeed) & vehSpeed <= config.Thresholds.CreepSpeed_kmh;

em1DriveEnergy = RCA_TrapzFinite(t, max(em1ElecPwr, 0)) / 3600;
em2DriveEnergy = RCA_TrapzFinite(t, max(em2ElecPwr, 0)) / 3600;
em1RegenEnergy = RCA_TrapzFinite(t, max(-em1ElecPwr, 0)) / 3600;
em2RegenEnergy = RCA_TrapzFinite(t, max(-em2ElecPwr, 0)) / 3600;
driveEnergy = RCA_TrapzFinite(t, max(totalElecPwr, 0)) / 3600;
regenEnergy = RCA_TrapzFinite(t, max(-totalElecPwr, 0)) / 3600;
mechDriveEnergy = RCA_TrapzFinite(t, max(totalMechPwr, 0)) / 3600;
lossEnergy = RCA_TrapzFinite(t, max(totalLossPwr, 0)) / 3600;

rows = RCA_AddKPI(rows, 'Electric Drive Traction Energy', driveEnergy, 'kWh', ...
    'Energy', 'Electric Drive', 'emot1_pwr + emot2_pwr', ...
    'Integrated positive electrical power across both machines.');
rows = RCA_AddKPI(rows, 'Electric Drive Regen Energy', regenEnergy, 'kWh', ...
    'Energy', 'Electric Drive', 'emot1_pwr + emot2_pwr', ...
    'Integrated electrical energy recovered during negative-power operation.');
rows = RCA_AddKPI(rows, 'Mechanical Output Energy', mechDriveEnergy, 'kWh', ...
    'Energy', 'Electric Drive', 'motor torque + motor speed', ...
    'Integrated positive shaft mechanical energy from both machines.');
rows = RCA_AddKPI(rows, 'Electric Drive Loss Energy', lossEnergy, 'kWh', ...
    'Losses', 'Electric Drive', 'emot1_loss_pwr + emot2_loss_pwr', ...
    'Integrated positive electric drive loss power.');
rows = RCA_AddKPI(rows, 'Average Tractive Efficiency', 100 * mean(tractiveEfficiency, 'omitnan'), '%', ...
    'Efficiency', 'Electric Drive', 'total electrical + mechanical power', ...
    'Computed only during positive drive-power operation.');
rows = RCA_AddKPI(rows, 'Average Regen Conversion Efficiency', 100 * mean(regenEfficiency, 'omitnan'), '%', ...
    'Efficiency', 'Electric Drive', 'total electrical + mechanical power', ...
    'Computed only during recuperation operation where shaft power is negative and electrical recovery is negative.');
rows = RCA_AddKPI(rows, 'High Loss Share in Drive', 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, driveMask), '%', ...
    'Losses', 'Electric Drive', 'loss power / electrical power', ...
    sprintf('Drive samples above %.1f%% instantaneous loss share are flagged as high-loss operation.', config.Thresholds.HighLossShare_pct));
rows = RCA_AddKPI(rows, 'Near Drive Limit Share', 100 * RCA_FractionTrue(nearDriveLimitMask, validPosLimit), '%', ...
    'Capability', 'Electric Drive', 'actual torque + max available torque', ...
    sprintf('Actual torque above %.0f%% of available positive torque is treated as near-limit drive operation.', config.Thresholds.LimitUsageFraction * 100));
rows = RCA_AddKPI(rows, 'Near Regen Limit Share', 100 * RCA_FractionTrue(nearRegenLimitMask, validNegLimit), '%', ...
    'Capability', 'Electric Drive', 'actual torque + min available torque', ...
    sprintf('Actual torque above %.0f%% of available recuperation torque magnitude is treated as near-limit regen operation.', config.Thresholds.LimitUsageFraction * 100));
rows = RCA_AddKPI(rows, 'Mean Motor Speed Magnitude', mean([em1SpdRpm; em2SpdRpm], 'omitnan'), 'rpm', ...
    'Operation', 'Electric Drive', 'emot1_act_spd + emot2_act_spd', ...
    'Average machine speed magnitude across both machines.');
rows = RCA_AddKPI(rows, 'High Motor Speed Time Share', 100 * RCA_FractionTrue(em1HighSpeedMask | em2HighSpeedMask, isfinite(em1SpdRpm) | isfinite(em2SpdRpm)), '%', ...
    'Operation', 'Electric Drive', 'motor speed', ...
    sprintf('High speed is defined as at least %.0f%% of the observed maximum motor speed.', config.Thresholds.HighMotorEfficiencySpeedFraction * 100));
rows = RCA_AddKPI(rows, 'Low-Speed High-Torque Share', 100 * RCA_FractionTrue(lowSpeedMask & abs(d.torqueActualTotal_Nm) > 0.5 * max(abs(d.torqueActualTotal_Nm), [], 'omitnan'), lowSpeedMask), '%', ...
    'Operation', 'Electric Drive', 'vehicle speed + motor torque', ...
    'Shows how much of low vehicle-speed operation occurs with high total electric-drive torque.');

rows = RCA_AddKPI(rows, 'Motor 1 Drive Energy', em1DriveEnergy, 'kWh', ...
    'Per Motor', 'Electric Drive', 'emot1_pwr', ...
    'Integrated positive electrical power for Motor 1.');
rows = RCA_AddKPI(rows, 'Motor 2 Drive Energy', em2DriveEnergy, 'kWh', ...
    'Per Motor', 'Electric Drive', 'emot2_pwr', ...
    'Integrated positive electrical power for Motor 2.');
rows = RCA_AddKPI(rows, 'Motor 1 Regen Energy', em1RegenEnergy, 'kWh', ...
    'Per Motor', 'Electric Drive', 'emot1_pwr', ...
    'Integrated electrical recovery for Motor 1.');
rows = RCA_AddKPI(rows, 'Motor 2 Regen Energy', em2RegenEnergy, 'kWh', ...
    'Per Motor', 'Electric Drive', 'emot2_pwr', ...
    'Integrated electrical recovery for Motor 2.');
rows = RCA_AddKPI(rows, 'Motor 1 Mean Drive Efficiency', 100 * mean(em1EfficiencyDrive, 'omitnan'), '%', ...
    'Per Motor', 'Electric Drive', 'Motor 1 electrical + mechanical power', ...
    'Computed only while Motor 1 is driving.');
rows = RCA_AddKPI(rows, 'Motor 2 Mean Drive Efficiency', 100 * mean(em2EfficiencyDrive, 'omitnan'), '%', ...
    'Per Motor', 'Electric Drive', 'Motor 2 electrical + mechanical power', ...
    'Computed only while Motor 2 is driving.');
rows = RCA_AddKPI(rows, 'Motor 1 Mean Regen Efficiency', 100 * mean(em1EfficiencyRegen, 'omitnan'), '%', ...
    'Per Motor', 'Electric Drive', 'Motor 1 electrical + mechanical power', ...
    'Computed only while Motor 1 is recuperating.');
rows = RCA_AddKPI(rows, 'Motor 2 Mean Regen Efficiency', 100 * mean(em2EfficiencyRegen, 'omitnan'), '%', ...
    'Per Motor', 'Electric Drive', 'Motor 2 electrical + mechanical power', ...
    'Computed only while Motor 2 is recuperating.');
rows = RCA_AddKPI(rows, 'Motor 1 Near Drive Limit Share', em1NearDrive, '%', ...
    'Per Motor', 'Electric Drive', 'emot1_act_trq + emot1_max_av_trq', ...
    'Share of Motor 1 drive operation close to available positive torque.');
rows = RCA_AddKPI(rows, 'Motor 2 Near Drive Limit Share', em2NearDrive, '%', ...
    'Per Motor', 'Electric Drive', 'emot2_act_trq + emot2_max_av_trq', ...
    'Share of Motor 2 drive operation close to available positive torque.');
rows = RCA_AddKPI(rows, 'Motor 1 Near Regen Limit Share', em1NearRegen, '%', ...
    'Per Motor', 'Electric Drive', 'emot1_act_trq + emot1_min_av_trq', ...
    'Share of Motor 1 regen operation close to available negative torque.');
rows = RCA_AddKPI(rows, 'Motor 2 Near Regen Limit Share', em2NearRegen, '%', ...
    'Per Motor', 'Electric Drive', 'emot2_act_trq + emot2_min_av_trq', ...
    'Share of Motor 2 regen operation close to available negative torque.');
rows = RCA_AddKPI(rows, 'Motor 1 Mean Drive Torque Reserve', mean(em1TorqueReserveDrive(em1DriveMask & isfinite(em1Max)), 'omitnan'), 'Nm', ...
    'Per Motor', 'Electric Drive', 'emot1_act_trq + emot1_max_av_trq', ...
    'Average positive torque headroom of Motor 1 during drive operation.');
rows = RCA_AddKPI(rows, 'Motor 2 Mean Drive Torque Reserve', mean(em2TorqueReserveDrive(em2DriveMask & isfinite(em2Max)), 'omitnan'), 'Nm', ...
    'Per Motor', 'Electric Drive', 'emot2_act_trq + emot2_max_av_trq', ...
    'Average positive torque headroom of Motor 2 during drive operation.');
rows = RCA_AddKPI(rows, 'Motor 1 Mean Regen Torque Reserve', mean(em1TorqueReserveRegen(em1RegenMask & isfinite(em1Min)), 'omitnan'), 'Nm', ...
    'Per Motor', 'Electric Drive', 'emot1_act_trq + emot1_min_av_trq', ...
    'Average recuperation torque headroom of Motor 1 during regen operation.');
rows = RCA_AddKPI(rows, 'Motor 2 Mean Regen Torque Reserve', mean(em2TorqueReserveRegen(em2RegenMask & isfinite(em2Min)), 'omitnan'), 'Nm', ...
    'Per Motor', 'Electric Drive', 'emot2_act_trq + emot2_min_av_trq', ...
    'Average recuperation torque headroom of Motor 2 during regen operation.');

summary(end + 1) = sprintf(['Electric drive converts %.2f kWh of electrical traction energy into %.2f kWh of mechanical output, ', ...
    'with %.2f kWh lost in the drive path. Mean tractive efficiency is %.1f%%.'], ...
    driveEnergy, mechDriveEnergy, lossEnergy, 100 * mean(tractiveEfficiency, 'omitnan'));
summary(end + 1) = sprintf(['Recuperation behavior: %.2f kWh of electrical recovery was observed with mean regen conversion efficiency of %.1f%%. ', ...
    'This indicates how effectively shaft braking power is converted back to electrical energy.'], ...
    regenEnergy, 100 * mean(regenEfficiency, 'omitnan'));
summary(end + 1) = sprintf(['Capability usage: near drive limit share is %.1f%% and near regen limit share is %.1f%%. ', ...
    'High values indicate the machines frequently operate near their torque envelope.'], ...
    100 * RCA_FractionTrue(nearDriveLimitMask, validPosLimit), 100 * RCA_FractionTrue(nearRegenLimitMask, validNegLimit));
summary(end + 1) = sprintf(['Per-machine context: Motor 1 drive efficiency is %.1f%% and Motor 2 drive efficiency is %.1f%%. ', ...
    'Large separation suggests unequal loading, thermal state differences, or calibration imbalance.'], ...
    100 * mean(em1EfficiencyDrive, 'omitnan'), 100 * mean(em2EfficiencyDrive, 'omitnan'));

if 100 * mean(tractiveEfficiency, 'omitnan') < 85
    recs(end + 1) = "Review motor and inverter operating-point clustering; the electric drive spends too much drive time away from an efficient speed-torque region.";
    evidence(end + 1) = sprintf('Average tractive efficiency is %.1f%%.', 100 * mean(tractiveEfficiency, 'omitnan'));
end
if 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, driveMask) > 15
    recs(end + 1) = "Investigate high-loss operating regions, thermal derating, or control allocation bias because instantaneous electric-drive losses remain elevated for a material share of drive operation.";
    evidence(end + 1) = sprintf('High-loss drive share is %.1f%% above the %.1f%% loss-share threshold.', ...
        100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, driveMask), config.Thresholds.HighLossShare_pct);
end
if 100 * RCA_FractionTrue(nearDriveLimitMask, validPosLimit) > 20
    recs(end + 1) = "Separate vehicle underperformance caused by machine torque saturation from issues upstream in the controller; the electric drive often operates close to its positive torque limit.";
    evidence(end + 1) = sprintf('Near drive limit share is %.1f%%.', 100 * RCA_FractionTrue(nearDriveLimitMask, validPosLimit));
end
if 100 * RCA_FractionTrue(nearRegenLimitMask, validNegLimit) > 20
    recs(end + 1) = "Review recuperation capability usage and blending strategy because the machines frequently run near their negative torque limit during regen opportunities.";
    evidence(end + 1) = sprintf('Near regen limit share is %.1f%%.', 100 * RCA_FractionTrue(nearRegenLimitMask, validNegLimit));
end
if abs(mean(em1EfficiencyDrive, 'omitnan') - mean(em2EfficiencyDrive, 'omitnan')) * 100 > 5
    recs(end + 1) = "Review Motor 1 and Motor 2 loading symmetry. A meaningful efficiency split between the two machines can indicate uneven torque sharing, hardware asymmetry, or inverter calibration bias.";
    evidence(end + 1) = sprintf('Motor 1 drive efficiency is %.1f%% and Motor 2 drive efficiency is %.1f%%.', ...
        100 * mean(em1EfficiencyDrive, 'omitnan'), 100 * mean(em2EfficiencyDrive, 'omitnan'));
end
if 100 * RCA_FractionTrue(em1HighSpeedMask | em2HighSpeedMask, isfinite(em1SpdRpm) | isfinite(em2SpdRpm)) > 15
    recs(end + 1) = "Investigate whether transmission usage or controller allocation keeps the machines at very high speed too often, which can elevate switching and mechanical losses.";
    evidence(end + 1) = sprintf('High motor-speed share is %.1f%%.', ...
        100 * RCA_FractionTrue(em1HighSpeedMask | em2HighSpeedMask, isfinite(em1SpdRpm) | isfinite(em2SpdRpm)));
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'ElectricDrive');
plotFiles = localAppendPlotFile(plotFiles, localPlotPowerOverview(figureFolder, t, totalElecPwr, totalMechPwr, totalLossPwr, em1SpdRpm, em2SpdRpm, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotLimitsAndTorque(figureFolder, t, em1Trq, em2Trq, em1Max, em2Max, em1Min, em2Min, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotOperatingMaps(figureFolder, em1SpdRpm, em1Trq, em1EfficiencyDrive, em1EfficiencyRegen, em2SpdRpm, em2Trq, em2EfficiencyDrive, em2EfficiencyRegen, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotRegenAndLoss(figureFolder, t, tractiveEfficiency, regenEfficiency, lossShare, driveLimitUsage, regenLimitUsage, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Electric Drive", recs, evidence);
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

function efficiency = localEfficiency(mechPwr, elecPwr, validMask)
efficiency = NaN(size(mechPwr));
efficiency(validMask) = mechPwr(validMask) ./ max(elecPwr(validMask), eps);
end

function efficiency = localRegenEfficiency(mechPwr, elecPwr, validMask)
efficiency = NaN(size(mechPwr));
efficiency(validMask) = abs(elecPwr(validMask)) ./ max(abs(mechPwr(validMask)), eps);
end

function sharePct = localMotorNearLimit(actualTorque, limitSignal, activeMask, usageFraction, positiveMode)
sharePct = NaN;
if positiveMode
    valid = activeMask & isfinite(actualTorque) & isfinite(limitSignal) & limitSignal > 0;
    sharePct = 100 * RCA_FractionTrue(actualTorque >= usageFraction .* limitSignal, valid);
else
    valid = activeMask & isfinite(actualTorque) & isfinite(limitSignal) & abs(limitSignal) > 0;
    sharePct = 100 * RCA_FractionTrue(abs(actualTorque) >= usageFraction .* abs(limitSignal), valid);
end
end

function plotFile = localPlotPowerOverview(outputFolder, t, elecPwr, mechPwr, lossPwr, em1SpdRpm, em2SpdRpm, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, elecPwr, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, mechPwr, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, lossPwr, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Electric Drive Power Flows');
ylabel('Power (kW)');
legend({'Electrical power', 'Mechanical power', 'Loss power', 'Zero line'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, em1SpdRpm, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, em2SpdRpm, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
title('Motor Speed Magnitude');
ylabel('Speed (rpm)');
legend({'Motor 1', 'Motor 2'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
cumElec = RCA_CumtrapzFinite(t, max(elecPwr, 0)) / 3600;
cumMech = RCA_CumtrapzFinite(t, max(mechPwr, 0)) / 3600;
cumLoss = RCA_CumtrapzFinite(t, max(lossPwr, 0)) / 3600;
plot(t, cumElec, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, cumMech, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, cumLoss, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Cumulative Electric Drive Energy');
xlabel('Time (s)');
ylabel('Energy (kWh)');
legend({'Electrical input', 'Mechanical output', 'Loss energy'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDrive_PowerOverview', config));
close(fig);
end

function plotFile = localPlotLimitsAndTorque(outputFolder, t, em1Trq, em2Trq, em1Max, em2Max, em1Min, em2Min, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 1, 1);
plot(t, em1Trq, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, em1Max, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, em1Min, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Motor 1 Torque and Available Limits');
ylabel('Torque (Nm)');
legend({'Actual torque', 'Max available', 'Min available', 'Zero line'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
plot(t, em2Trq, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, em2Max, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, em2Min, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Motor 2 Torque and Available Limits');
xlabel('Time (s)');
ylabel('Torque (Nm)');
legend({'Actual torque', 'Max available', 'Min available', 'Zero line'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDrive_TorqueAndLimits', config));
close(fig);
end

function plotFile = localPlotOperatingMaps(outputFolder, em1SpdRpm, em1Trq, em1EffDrive, em1EffRegen, em2SpdRpm, em2Trq, em2EffDrive, em2EffRegen, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

em1EffCombined = em1EffDrive;
em1EffCombined(isfinite(em1EffRegen)) = em1EffRegen(isfinite(em1EffRegen));
em2EffCombined = em2EffDrive;
em2EffCombined(isfinite(em2EffRegen)) = em2EffRegen(isfinite(em2EffRegen));

subplot(2, 1, 1);
valid1 = isfinite(em1SpdRpm) & isfinite(em1Trq) & isfinite(em1EffCombined);
scatter(em1SpdRpm(valid1), em1Trq(valid1), 12, em1EffCombined(valid1) * 100, 'filled'); hold on;
yline(0, '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Motor 1 Operating Map');
xlabel('Motor speed (rpm)');
ylabel('Motor torque (Nm)');
cb1 = colorbar;
cb1.Label.String = 'Drive / regen efficiency (%)';
grid on;

subplot(2, 1, 2);
valid2 = isfinite(em2SpdRpm) & isfinite(em2Trq) & isfinite(em2EffCombined);
scatter(em2SpdRpm(valid2), em2Trq(valid2), 12, em2EffCombined(valid2) * 100, 'filled'); hold on;
yline(0, '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Motor 2 Operating Map');
xlabel('Motor speed (rpm)');
ylabel('Motor torque (Nm)');
cb2 = colorbar;
cb2.Label.String = 'Drive / regen efficiency (%)';
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDrive_OperatingMaps', config));
close(fig);
end

function plotFile = localPlotRegenAndLoss(outputFolder, t, tractiveEfficiency, regenEfficiency, lossShare, driveLimitUsage, regenLimitUsage, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
plot(t, tractiveEfficiency * 100, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
title('Tractive Efficiency');
ylabel('Efficiency (%)');
grid on;

subplot(2, 2, 2);
plot(t, regenEfficiency * 100, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Regen Conversion Efficiency');
ylabel('Efficiency (%)');
grid on;

subplot(2, 2, 3);
plot(t, lossShare, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.HighLossShare_pct, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Instantaneous Loss Share in Drive');
xlabel('Time (s)');
ylabel('Loss share (%)');
legend({'Loss share', 'High-loss threshold'}, 'Location', 'best');
grid on;

subplot(2, 2, 4);
plot(t, driveLimitUsage, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, regenLimitUsage, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.LimitUsageFraction * 100, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Torque Limit Usage');
xlabel('Time (s)');
ylabel('Limit usage (%)');
legend({'Drive limit usage', 'Regen limit usage', 'Near-limit threshold'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDrive_RegenAndLoss', config));
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
