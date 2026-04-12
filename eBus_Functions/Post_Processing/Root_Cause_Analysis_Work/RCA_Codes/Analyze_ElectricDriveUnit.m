function result = Analyze_ElectricDriveUnit(analysisData, outputPaths, config)
% Analyze_ElectricDriveUnit  Integrated motor, gearbox, and final-drive RCA.
%
% The Electric Drive Unit (EDU) combines the electric machines/inverters,
% gearbox/transmission, and final-drive force path into one subsystem-level
% RCA owner view. The purpose is to evaluate the complete conversion chain:
% electrical power -> motor shaft torque -> gearbox output -> tractive force.

result = localInitResult("ELECTRIC DRIVE UNIT", ...
    {'emot1_act_trq', 'emot2_act_trq', 'emot1_act_spd', 'emot2_act_spd', 'emot1_pwr', 'emot2_pwr', 'gbx_out_trq', 'net_trac_trq'}, ...
    {'emot1_loss_pwr', 'emot2_loss_pwr', 'gbx_pwr_loss', 'gbx_out_spd', 'gr_num', 'gr_ratio', 'veh_long_force', 'whl_force'});

d = analysisData.Derived;
t = d.time_s(:);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Electric Drive Unit analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.SummaryText = summary;
    result.Suggestions = RCA_MakeSuggestionTable("Electric Drive Unit", recs, evidence);
    return;
end

unitFolder = fullfile(outputPaths.FiguresSubsystem, 'ElectricDriveUnit');
if ~exist(unitFolder, 'dir')
    mkdir(unitFolder);
end

% Run legacy module analyses internally so the combined EDU result preserves
% mature motor, transmission, and final-drive KPIs/plots without exposing
% three separate subsystem owners in the top-level RCA flow.
subResults = [ ...
    localRunLegacySubmodule(@Analyze_ElectricDrive, "Electric Drive", analysisData, outputPaths, config); ...
    localRunLegacySubmodule(@Analyze_Transmission, "Transmission", analysisData, outputPaths, config); ...
    localRunLegacySubmodule(@Analyze_FinalDrive, "Final Drive", analysisData, outputPaths, config)];

motorElecPower = d.motorElectricalPower_kW(:);
motorMechPower = d.motorMechanicalPower_kW(:);
motorLossPower = d.motorLossPower_kW(:);
gearboxLossPower = d.gearboxLossPower_kW(:);
gbxOutTrq = d.gearboxOutputTorque_Nm(:);
gbxOutSpd = d.gearboxOutputSpeed_rads(:);
tractionForce = d.tractionForce_N(:);
tractionPower = d.tractionPower_kW(:);
gear = d.gearNumber(:);
gearRatio = d.gearRatio(:);
motorSpeed = d.motorSpeed_rpm(:);
motorTorque = d.torqueActualTotal_Nm(:);
torqueDemand = d.torqueDemandTotal_Nm(:);
posTorqueLimit = d.torquePositiveLimit_Nm(:);
negTorqueLimit = d.torqueNegativeLimit_Nm(:);
eduEvidenceAvailable = any(isfinite(motorElecPower)) || any(isfinite(motorMechPower)) || ...
    any(isfinite(gbxOutTrq)) || any(isfinite(tractionForce));
em1Trq = localSignalVector(analysisData, 'emot1_act_trq', numel(t));
em2Trq = localSignalVector(analysisData, 'emot2_act_trq', numel(t));
em1Spd = localSignalVector(analysisData, 'emot1_act_spd', numel(t));
em2Spd = localSignalVector(analysisData, 'emot2_act_spd', numel(t));
em1Loss = localSignalVector(analysisData, 'emot1_loss_pwr', numel(t));
em2Loss = localSignalVector(analysisData, 'emot2_loss_pwr', numel(t));

gearboxOutputPower = gbxOutTrq .* gbxOutSpd / 1000;
if ~any(isfinite(gearboxOutputPower))
    gearboxOutputPower = tractionPower + gearboxLossPower;
end

elecDriveEnergy = RCA_TrapzFinite(t, max(motorElecPower, 0)) / 3600;
elecRegenEnergy = RCA_TrapzFinite(t, max(-motorElecPower, 0)) / 3600;
motorMechDriveEnergy = RCA_TrapzFinite(t, max(motorMechPower, 0)) / 3600;
gbxDriveEnergy = RCA_TrapzFinite(t, max(gearboxOutputPower, 0)) / 3600;
roadDriveEnergy = RCA_TrapzFinite(t, max(tractionPower, 0)) / 3600;
roadRegenEnergy = RCA_TrapzFinite(t, max(-tractionPower, 0)) / 3600;
motorLossEnergy = RCA_TrapzFinite(t, max(motorLossPower, 0)) / 3600;
gearboxLossEnergy = RCA_TrapzFinite(t, max(gearboxLossPower, 0)) / 3600;
unitLossEnergy = motorLossEnergy + gearboxLossEnergy;

driveEfficiency = 100 * roadDriveEnergy / max(elecDriveEnergy, eps);
shaftToRoadEfficiency = 100 * roadDriveEnergy / max(motorMechDriveEnergy, eps);
regenRecovery = 100 * elecRegenEnergy / max(roadRegenEnergy, eps);
unitLossShare = 100 * unitLossEnergy / max(elecDriveEnergy, eps);

driveMask = isfinite(motorTorque) & motorTorque > config.Thresholds.ControllerTorqueDeadband_Nm;
regenMask = isfinite(motorTorque) & motorTorque < -config.Thresholds.ControllerTorqueDeadband_Nm;
validPosLimit = driveMask & isfinite(posTorqueLimit) & posTorqueLimit > 0;
validNegLimit = regenMask & isfinite(negTorqueLimit) & abs(negTorqueLimit) > 0;
nearPositiveLimit = validPosLimit & motorTorque >= config.Thresholds.LimitUsageFraction .* posTorqueLimit;
nearRegenLimit = validNegLimit & abs(motorTorque) >= config.Thresholds.LimitUsageFraction .* abs(negTorqueLimit);
feasibleDemand = localClampDemandToAvailable(torqueDemand, posTorqueLimit, negTorqueLimit);
validDemandActual = isfinite(torqueDemand) & isfinite(motorTorque);
validFeasibleActual = isfinite(feasibleDemand) & isfinite(motorTorque);
torqueTrackingError = motorTorque - torqueDemand;
feasibleTrackingError = motorTorque - feasibleDemand;
driveDemandMask = isfinite(torqueDemand) & torqueDemand > config.Thresholds.ControllerTorqueDeadband_Nm;
regenDemandMask = isfinite(torqueDemand) & torqueDemand < -config.Thresholds.ControllerTorqueDeadband_Nm;
demandPositiveLimited = driveDemandMask & isfinite(posTorqueLimit) & torqueDemand > posTorqueLimit + config.Thresholds.ControllerTorqueDeadband_Nm;
demandRegenLimited = regenDemandMask & isfinite(negTorqueLimit) & abs(torqueDemand) > abs(negTorqueLimit) + config.Thresholds.ControllerTorqueDeadband_Nm;
withinCapability = validDemandActual & ~demandPositiveLimited & ~demandRegenLimited;
fullyMetMask = validDemandActual & abs(torqueTrackingError) <= config.Thresholds.ControllerTorqueTrackingWarn_Nm;
driveShortfall = max(torqueDemand - motorTorque, 0);
regenShortfall = max(abs(torqueDemand) - abs(motorTorque), 0);
driveShortfall(~driveDemandMask) = NaN;
regenShortfall(~regenDemandMask) = NaN;
positiveReserve = posTorqueLimit - torqueDemand;
positiveReserve(~driveDemandMask | ~isfinite(posTorqueLimit)) = NaN;
regenReserve = abs(negTorqueLimit) - abs(torqueDemand);
regenReserve(~regenDemandMask | ~isfinite(negTorqueLimit)) = NaN;

shiftMask = localShiftMask(gear);
shiftEvents = localShiftEvents(gear);
shiftIdx = find(shiftEvents);
shiftCount = sum(shiftEvents);
lossPowerTotal = max(motorLossPower, 0) + max(gearboxLossPower, 0);
driveLossShare = NaN(size(t));
drivePowerPositive = max(motorElecPower, 0);
driveLossShare(drivePowerPositive > eps) = 100 * lossPowerTotal(drivePowerPositive > eps) ./ ...
    max(drivePowerPositive(drivePowerPositive > eps), eps);
shiftTorqueDipPct = localShiftTorqueDip(t, gbxOutTrq, shiftIdx);
shiftLossSpike_kW = localShiftLossSpike(t, lossPowerTotal, shiftIdx);
propulsionSignal = localPropulsionDeliverySignal(d, tractionForce);
propulsionSlew = localSlewRate(t, propulsionSignal);
propulsionSmoothness = sqrt(mean(propulsionSlew(isfinite(propulsionSlew)).^2, 'omitnan'));
propulsionRipplePct = localRipplePct(propulsionSignal, driveMask);
torqueImbalancePct = localBalanceImbalancePct(em1Trq, em2Trq);
lossImbalancePct = localBalanceImbalancePct(em1Loss, em2Loss);
speedMismatchRpm = abs(em1Spd - em2Spd) * 60 / (2 * pi);

forceConsistencyResidual = NaN(size(t));
if any(isfinite(tractionForce)) && any(isfinite(gbxOutTrq)) && isfinite(d.finalDriveRatio) && isfinite(d.wheelRadius_m) && d.wheelRadius_m > eps
    estimatedForce = gbxOutTrq .* d.finalDriveRatio ./ d.wheelRadius_m;
    forceConsistencyResidual = tractionForce - estimatedForce;
end

rows = RCA_AddKPI(rows, 'EDU Electrical Drive Energy', elecDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'emot1_pwr + emot2_pwr', ...
    'Electrical drive-positive energy entering the electric machines and inverters.');
rows = RCA_AddKPI(rows, 'EDU Electrical Regen Energy', elecRegenEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'emot1_pwr + emot2_pwr', ...
    'Electrical recuperation energy returned from the EDU toward the HV system.');
rows = RCA_AddKPI(rows, 'EDU Motor Shaft Drive Energy', motorMechDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'motor torque + motor speed', ...
    'Positive mechanical energy at motor shafts before gearbox losses.');
rows = RCA_AddKPI(rows, 'EDU Gearbox Output Drive Energy', gbxDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'gbx_out_trq + gbx_out_spd', ...
    'Positive mechanical energy leaving the gearbox toward the final drive.');
rows = RCA_AddKPI(rows, 'EDU Road Tractive Drive Energy', roadDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'tractive force + vehicle speed', ...
    'Positive tractive energy delivered at the road interface.');
rows = RCA_AddKPI(rows, 'EDU Torque Demand Fully Met Time Share', 100 * RCA_FractionTrue(fullyMetMask, validDemandActual), '%', ...
    'Torque Delivery', 'Electric Drive Unit', 'torque demand + actual torque', ...
    sprintf('Demand is treated as met when demand-to-actual torque error is within %.0f Nm.', config.Thresholds.ControllerTorqueTrackingWarn_Nm));
rows = RCA_AddKPI(rows, 'EDU Torque Tracking MAE', mean(abs(torqueTrackingError(validDemandActual)), 'omitnan'), 'Nm', ...
    'Torque Delivery', 'Electric Drive Unit', 'torque demand + actual torque', ...
    'Mean absolute error between controller torque demand and total actual motor torque.');
rows = RCA_AddKPI(rows, 'EDU Torque Tracking RMSE', sqrt(mean(torqueTrackingError(validDemandActual).^2, 'omitnan')), 'Nm', ...
    'Torque Delivery', 'Electric Drive Unit', 'torque demand + actual torque', ...
    'RMS error between controller torque demand and total actual motor torque.');
rows = RCA_AddKPI(rows, 'EDU Feasible Torque Realization MAE', mean(abs(feasibleTrackingError(validFeasibleActual)), 'omitnan'), 'Nm', ...
    'Torque Delivery', 'Electric Drive Unit', 'torque demand + available limits + actual torque', ...
    'Mean absolute error after clipping demand to available drive and regen torque limits. High value points to realization dynamics rather than pure capability.');
rows = RCA_AddKPI(rows, 'EDU Demand Exceeds Positive Capability Time Share', 100 * RCA_FractionTrue(demandPositiveLimited, driveDemandMask), '%', ...
    'Capability', 'Electric Drive Unit', 'torque demand + max available torque', ...
    'Share of positive-demand operation where requested torque exceeds available positive torque.');
rows = RCA_AddKPI(rows, 'EDU Regen Demand Exceeds Capability Time Share', 100 * RCA_FractionTrue(demandRegenLimited, regenDemandMask), '%', ...
    'Capability', 'Electric Drive Unit', 'torque demand + min available torque', ...
    'Share of negative-demand operation where requested recuperation torque exceeds available negative torque.');
rows = RCA_AddKPI(rows, 'EDU Peak Drive Torque Shortfall', max(driveShortfall, [], 'omitnan'), 'Nm', ...
    'Torque Delivery', 'Electric Drive Unit', 'positive torque demand + actual torque', ...
    'Peak positive torque delivery shortfall during drive demand.');
rows = RCA_AddKPI(rows, 'EDU Peak Regen Torque Shortfall', max(regenShortfall, [], 'omitnan'), 'Nm', ...
    'Torque Delivery', 'Electric Drive Unit', 'negative torque demand + actual torque', ...
    'Peak recuperation torque delivery shortfall during regen demand.');
rows = RCA_AddKPI(rows, 'EDU Mean Positive Torque Reserve', mean(positiveReserve, 'omitnan'), 'Nm', ...
    'Capability', 'Electric Drive Unit', 'max available torque - positive demand torque', ...
    'Average propulsion torque headroom during positive torque demand.');
rows = RCA_AddKPI(rows, 'EDU Mean Regen Torque Reserve', mean(regenReserve, 'omitnan'), 'Nm', ...
    'Capability', 'Electric Drive Unit', 'min available torque - negative demand torque', ...
    'Average recuperation torque headroom during negative torque demand.');
rows = RCA_AddKPI(rows, 'EDU Motor/Inverter Loss Energy', motorLossEnergy, 'kWh', ...
    'Losses', 'Electric Drive Unit', 'emot1_loss_pwr + emot2_loss_pwr', ...
    'Integrated motor and inverter loss energy.');
rows = RCA_AddKPI(rows, 'EDU Gearbox Loss Energy', gearboxLossEnergy, 'kWh', ...
    'Losses', 'Electric Drive Unit', 'gbx_pwr_loss', ...
    'Integrated gearbox/transmission loss energy.');
rows = RCA_AddKPI(rows, 'EDU Total Internal Loss Energy', unitLossEnergy, 'kWh', ...
    'Losses', 'Electric Drive Unit', 'motor loss + gearbox loss', ...
    'Motor/inverter plus gearbox loss energy inside the combined Electric Drive Unit.');
rows = RCA_AddKPI(rows, 'EDU Electrical-to-Road Drive Efficiency', driveEfficiency, '%', ...
    'Efficiency', 'Electric Drive Unit', 'motor electrical energy + road tractive energy', ...
    'Approximate drive efficiency from motor electrical power to road tractive power.');
rows = RCA_AddKPI(rows, 'EDU Shaft-to-Road Drive Efficiency', shaftToRoadEfficiency, '%', ...
    'Efficiency', 'Electric Drive Unit', 'motor shaft energy + road tractive energy', ...
    'Approximate mechanical transfer efficiency from motor shafts through gearbox/final-drive path.');
rows = RCA_AddKPI(rows, 'EDU Road-to-Electrical Regen Recovery', regenRecovery, '%', ...
    'Efficiency', 'Electric Drive Unit', 'road regen energy + motor electrical regen', ...
    'Approximate recuperation recovery from road negative work to electrical regen energy.');
rows = RCA_AddKPI(rows, 'EDU Loss Share of Electrical Drive Energy', unitLossShare, '%', ...
    'Losses', 'Electric Drive Unit', 'loss energy / electrical drive energy', ...
    'Combined motor/inverter and gearbox loss share normalized by drive electrical energy.');
rows = RCA_AddKPI(rows, 'EDU Near Positive Torque Limit Share', 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit), '%', ...
    'Capability', 'Electric Drive Unit', 'actual torque + max available torque', ...
    sprintf('Actual positive torque above %.0f%% of available torque indicates propulsion envelope usage.', 100 * config.Thresholds.LimitUsageFraction));
rows = RCA_AddKPI(rows, 'EDU Near Regen Torque Limit Share', 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit), '%', ...
    'Capability', 'Electric Drive Unit', 'actual torque + min available torque', ...
    sprintf('Actual negative torque above %.0f%% of available regen torque indicates recuperation envelope usage.', 100 * config.Thresholds.LimitUsageFraction));
rows = RCA_AddKPI(rows, 'EDU Shift Count', shiftCount, 'count', ...
    'Gear Operation', 'Electric Drive Unit', 'gr_num', ...
    'Actual gear changes inside the combined unit.');
rows = RCA_AddKPI(rows, 'EDU Mean Shift Torque Dip', mean(shiftTorqueDipPct, 'omitnan'), '%', ...
    'Gear Operation', 'Electric Drive Unit', 'gr_num + gbx_out_trq', ...
    'Mean gearbox output torque dip around detected gear changes. High values indicate shift torque interruption.');
rows = RCA_AddKPI(rows, 'EDU Peak Shift Torque Dip', max(shiftTorqueDipPct, [], 'omitnan'), '%', ...
    'Gear Operation', 'Electric Drive Unit', 'gr_num + gbx_out_trq', ...
    'Worst gearbox output torque dip around detected gear changes.');
rows = RCA_AddKPI(rows, 'EDU Mean Shift Loss Spike', mean(shiftLossSpike_kW, 'omitnan'), 'kW', ...
    'Gear Operation', 'Electric Drive Unit', 'gr_num + motor loss + gearbox loss', ...
    'Mean increase in EDU loss power near detected gear changes relative to local pre-shift baseline.');
rows = RCA_AddKPI(rows, 'EDU High Loss Share Time', 100 * RCA_FractionTrue(driveLossShare >= config.Thresholds.HighLossShare_pct, isfinite(driveLossShare)), '%', ...
    'Losses', 'Electric Drive Unit', 'motor loss + gearbox loss + motor electrical power', ...
    sprintf('Drive samples above %.1f%% internal loss share are flagged as high-loss EDU operation.', config.Thresholds.HighLossShare_pct));
rows = RCA_AddKPI(rows, 'EDU Force Path Residual MAE', mean(abs(forceConsistencyResidual), 'omitnan'), 'N', ...
    'Consistency', 'Electric Drive Unit', 'gbx_out_trq + final drive ratio + wheel radius + tractive force', ...
    'Checks whether gearbox torque, final-drive ratio, wheel radius, and tractive force are mutually consistent.');
rows = RCA_AddKPI(rows, 'EDU Propulsion Delivery Slew RMS', propulsionSmoothness, 'Nm/s or N/s', ...
    'Delivery Quality', 'Electric Drive Unit', 'final drive torque or tractive force + time', ...
    'RMS rate of change of delivered propulsion signal. High value indicates abrupt wheel-side torque delivery or shift disturbance.');
rows = RCA_AddKPI(rows, 'EDU Propulsion Delivery Ripple', propulsionRipplePct, '%', ...
    'Delivery Quality', 'Electric Drive Unit', 'final drive torque or tractive force', ...
    'Coefficient-of-variation style ripple metric for delivered propulsion during positive drive operation.');
rows = RCA_AddKPI(rows, 'EDU Mean Motor Torque Imbalance', mean(torqueImbalancePct, 'omitnan'), '%', ...
    'Motor Balance', 'Electric Drive Unit', 'emot1_act_trq + emot2_act_trq', ...
    'Mean torque-sharing imbalance between Motor 1 and Motor 2. High value can indicate torque allocation asymmetry.');
rows = RCA_AddKPI(rows, 'EDU Mean Motor Loss Imbalance', mean(lossImbalancePct, 'omitnan'), '%', ...
    'Motor Balance', 'Electric Drive Unit', 'emot1_loss_pwr + emot2_loss_pwr', ...
    'Mean loss-power imbalance between Motor 1 and Motor 2. High value can indicate unequal loading or hardware/calibration asymmetry.');
rows = RCA_AddKPI(rows, 'EDU Mean Motor Speed Mismatch', mean(speedMismatchRpm, 'omitnan'), 'rpm', ...
    'Motor Balance', 'Electric Drive Unit', 'emot1_act_spd + emot2_act_spd', ...
    'Mean speed mismatch between motors. Mechanically coupled machines should normally remain close unless architecture allows speed separation.');

summary(end + 1) = sprintf(['Electric Drive Unit energy path: %.2f kWh electrical drive energy enters the machines, ', ...
    '%.2f kWh reaches the road as positive tractive work, and %.2f kWh is recorded as motor plus gearbox loss.'], ...
    elecDriveEnergy, roadDriveEnergy, unitLossEnergy);
summary(end + 1) = sprintf(['Demand-capability-delivery chain: %.1f%% of valid demand samples are met within %.0f Nm. ', ...
    'Positive demand exceeds available drive torque for %.1f%% of drive demand, while regen demand exceeds available negative torque for %.1f%% of regen demand.'], ...
    100 * RCA_FractionTrue(fullyMetMask, validDemandActual), config.Thresholds.ControllerTorqueTrackingWarn_Nm, ...
    100 * RCA_FractionTrue(demandPositiveLimited, driveDemandMask), 100 * RCA_FractionTrue(demandRegenLimited, regenDemandMask));
summary(end + 1) = sprintf(['Root-cause separation logic: demand above available torque is treated as capability limitation; ', ...
    'demand inside available limits but high actual error is treated as realization/control or actuator dynamics; ', ...
    'good motor realization with poor road-side delivery points to gearbox/final-drive transfer; high loss share points to energy-conversion efficiency.']);
summary(end + 1) = sprintf(['Electric Drive Unit efficiency summary: electrical-to-road drive efficiency is %.1f%%, ', ...
    'shaft-to-road mechanical efficiency is %.1f%%, and road-to-electrical regen recovery is %.1f%%.'], ...
    driveEfficiency, shaftToRoadEfficiency, regenRecovery);
summary(end + 1) = sprintf(['Gear and capability context: %d shifts were observed; near positive torque limit share is %.1f%% ', ...
    'and near regen torque limit share is %.1f%%.'], ...
    shiftCount, 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit), 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit));
summary(end + 1) = sprintf(['Shift and delivery quality: mean shift torque dip is %.1f%%, peak shift torque dip is %.1f%%, ', ...
    'and propulsion delivery ripple during drive is %.1f%%.'], ...
    mean(shiftTorqueDipPct, 'omitnan'), max(shiftTorqueDipPct, [], 'omitnan'), propulsionRipplePct);
summary(end + 1) = sprintf(['Dual-motor balance: mean torque imbalance is %.1f%%, mean loss imbalance is %.1f%%, ', ...
    'and mean motor speed mismatch is %.1f rpm.'], ...
    mean(torqueImbalancePct, 'omitnan'), mean(lossImbalancePct, 'omitnan'), mean(speedMismatchRpm, 'omitnan'));
if isfinite(mean(abs(forceConsistencyResidual), 'omitnan'))
    summary(end + 1) = sprintf('Final-drive force-path consistency residual MAE is %.0f N, based on gearbox torque, final-drive ratio, wheel radius, and tractive force.', ...
        mean(abs(forceConsistencyResidual), 'omitnan'));
end

if unitLossShare > config.Thresholds.HighLossShare_pct
    recs(end + 1) = "Review the combined motor/inverter and gearbox loss path; EDU internal losses are high relative to electrical drive energy.";
    evidence(end + 1) = sprintf('EDU loss share of electrical drive energy is %.1f%%.', unitLossShare);
end
if 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit) > 20
    recs(end + 1) = "Separate upstream torque-command issues from EDU propulsion capability limits; the unit frequently operates near the positive torque envelope.";
    evidence(end + 1) = sprintf('Near positive torque limit share is %.1f%%.', 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit));
end
if 100 * RCA_FractionTrue(demandPositiveLimited, driveDemandMask) > 10
    recs(end + 1) = "Review EDU torque-speed envelope, ratio selection, and high-speed power-limited operation because controller demand frequently exceeds available positive torque.";
    evidence(end + 1) = sprintf('Positive demand exceeds available torque for %.1f%% of drive demand.', 100 * RCA_FractionTrue(demandPositiveLimited, driveDemandMask));
end
if 100 * RCA_FractionTrue(withinCapability & abs(feasibleTrackingError) > config.Thresholds.ControllerTorqueTrackingWarn_Nm, withinCapability) > 10
    recs(end + 1) = "Review torque realization dynamics, filtering, rate limits, or torque arbitration because demand is often feasible but not accurately delivered.";
    evidence(end + 1) = sprintf('Feasible torque realization MAE is %.1f Nm.', mean(abs(feasibleTrackingError(validFeasibleActual)), 'omitnan'));
end
if 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit) > 20
    recs(end + 1) = "Review recuperation torque envelope and gear-dependent regen capability; the EDU frequently operates near the negative torque limit.";
    evidence(end + 1) = sprintf('Near regen torque limit share is %.1f%%.', 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit));
end
if shiftCount > 0 && mean(lossPowerTotal(shiftMask), 'omitnan') > 1.25 * mean(lossPowerTotal(~shiftMask), 'omitnan')
    recs(end + 1) = "Review shift torque handover and ratio transition logic; EDU losses are materially higher during shift activity.";
    evidence(end + 1) = sprintf('Mean EDU loss power during shifts is %.2f kW versus %.2f kW away from shifts.', ...
        mean(lossPowerTotal(shiftMask), 'omitnan'), mean(lossPowerTotal(~shiftMask), 'omitnan'));
end
if mean(shiftTorqueDipPct, 'omitnan') > 15
    recs(end + 1) = "Review shift torque-fill and clutch/ratio transition calibration; gearbox output torque dips materially during shifts.";
    evidence(end + 1) = sprintf('Mean shift torque dip is %.1f%% and peak shift torque dip is %.1f%%.', mean(shiftTorqueDipPct, 'omitnan'), max(shiftTorqueDipPct, [], 'omitnan'));
end
if mean(torqueImbalancePct, 'omitnan') > 15 || mean(lossImbalancePct, 'omitnan') > 15
    recs(end + 1) = "Review Motor 1 and Motor 2 torque allocation and loss balance; persistent imbalance can reduce efficiency or hide one-machine saturation.";
    evidence(end + 1) = sprintf('Mean torque imbalance is %.1f%% and mean loss imbalance is %.1f%%.', mean(torqueImbalancePct, 'omitnan'), mean(lossImbalancePct, 'omitnan'));
end
if isfinite(mean(abs(forceConsistencyResidual), 'omitnan')) && mean(abs(forceConsistencyResidual), 'omitnan') > config.Thresholds.ForceBalanceResidualWarn_N
    recs(end + 1) = "Check final-drive ratio, wheel-radius specification, or logged tractive-force convention; force-path residual is above the configured review threshold.";
    evidence(end + 1) = sprintf('Force-path residual MAE is %.0f N.', mean(abs(forceConsistencyResidual), 'omitnan'));
end

plotFiles = localAppendPlotFile(plotFiles, localPlotDemandCapabilityDelivery(unitFolder, t, torqueDemand, feasibleDemand, motorTorque, posTorqueLimit, negTorqueLimit, driveShortfall, regenShortfall, propulsionSignal, gear, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotEnergyPath(unitFolder, elecDriveEnergy, motorMechDriveEnergy, gbxDriveEnergy, roadDriveEnergy, elecRegenEnergy, roadRegenEnergy, motorLossEnergy, gearboxLossEnergy, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotTorqueForcePath(unitFolder, t, motorTorque, gbxOutTrq, tractionForce, gear, gearRatio, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotLossAndShift(unitFolder, t, motorElecPower, motorLossPower, gearboxLossPower, driveLossShare, shiftMask, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotOperatingMap(unitFolder, motorSpeed, motorTorque, gear, driveLossShare, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotMotorBalance(unitFolder, t, em1Trq, em2Trq, em1Loss, em2Loss, speedMismatchRpm, torqueImbalancePct, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotShiftDisturbance(unitFolder, t, shiftIdx, shiftTorqueDipPct, shiftLossSpike_kW, gbxOutTrq, propulsionSignal, gear, config));

mergedKpis = RCA_FinalizeKPITable(rows);
mergedKpis = [mergedKpis; localMergeSubmoduleKPIs(subResults)];
mergedSuggestions = RCA_MakeSuggestionTable("Electric Drive Unit", recs, evidence);
mergedSuggestions = [mergedSuggestions; localMergeSubmoduleSuggestions(subResults)];

plotFiles = [plotFiles; localMergeSubmoduleFigures(subResults)];
plotFiles = localExistingFiles(plotFiles);

for iSub = 1:numel(subResults)
    if numel(string(subResults(iSub).SummaryText)) > 0
        legacySummary = "Submodule evidence - " + string(subResults(iSub).Name) + ": " + string(subResults(iSub).SummaryText(:));
        summary = [summary(:); legacySummary(:)]; %#ok<AGROW>
    end
    if numel(string(subResults(iSub).Warnings)) > 0
        legacyWarnings = string(subResults(iSub).Warnings(:));
        result.Warnings = [result.Warnings(:); legacyWarnings(:)]; %#ok<AGROW>
    end
end

result.Available = eduEvidenceAvailable || any([subResults.Available]);
result.KPITable = mergedKpis;
result.FigureFiles = plotFiles;
result.SummaryText = unique(summary(summary ~= ""));
result.Suggestions = mergedSuggestions;
end

function kpiTable = localMergeSubmoduleKPIs(subResults)
kpiTable = RCA_FinalizeKPITable([]);
for iSub = 1:numel(subResults)
    if ~istable(subResults(iSub).KPITable) || height(subResults(iSub).KPITable) == 0
        continue;
    end
    tableValue = subResults(iSub).KPITable;
    if ~ismember('Subsystem', tableValue.Properties.VariableNames)
        tableValue.Subsystem = repmat("Electric Drive Unit", height(tableValue), 1);
    end
    tableValue.Subsystem(:) = "Electric Drive Unit";
    tableValue.Category = string(subResults(iSub).Name) + " / " + string(tableValue.Category);
    kpiTable = [kpiTable; tableValue]; %#ok<AGROW>
end
end

function suggestionTable = localMergeSubmoduleSuggestions(subResults)
suggestionTable = RCA_MakeSuggestionTable("Electric Drive Unit", strings(0, 1), strings(0, 1));
for iSub = 1:numel(subResults)
    if ~isfield(subResults(iSub), 'Suggestions') || ~istable(subResults(iSub).Suggestions) || height(subResults(iSub).Suggestions) == 0
        continue;
    end
    tableValue = subResults(iSub).Suggestions;
    tableValue.Subsystem(:) = "Electric Drive Unit";
    tableValue.Recommendation = string(subResults(iSub).Name) + " evidence: " + string(tableValue.Recommendation);
    suggestionTable = [suggestionTable; tableValue]; %#ok<AGROW>
end
end

function figureFiles = localMergeSubmoduleFigures(subResults)
figureFiles = strings(0, 1);
for iSub = 1:numel(subResults)
    if isfield(subResults(iSub), 'FigureFiles')
        figureFiles = [figureFiles; string(subResults(iSub).FigureFiles(:))]; %#ok<AGROW>
    end
end
end

function files = localExistingFiles(files)
files = string(files(:));
files = files(files ~= "");
keep = false(size(files));
for iFile = 1:numel(files)
    keep(iFile) = isfile(char(files(iFile)));
end
files = files(keep);
end

function shiftEvents = localShiftEvents(gear)
shiftEvents = false(size(gear));
if numel(gear) < 2
    return;
end
valid = isfinite(gear(2:end)) & isfinite(gear(1:end - 1));
shiftEvents(2:end) = valid & abs(diff(gear)) > 0.05;
end

function shiftMask = localShiftMask(gear)
shiftMask = localShiftEvents(gear);
if any(shiftMask)
    shiftMask = shiftMask | [shiftMask(2:end); false] | [false; shiftMask(1:end - 1)];
end
end

function signalData = localSignalVector(analysisData, signalName, n)
signalData = NaN(n, 1);
if nargin < 3 || n <= 0 || ~isstruct(analysisData) || ~isfield(analysisData, 'Signals')
    return;
end
try
    entry = RCA_GetSignalData(analysisData.Signals, signalName);
    if entry.Available && isnumeric(entry.Data)
        signalData = localColumnVector(entry.Data, n);
    end
catch
    signalData = NaN(n, 1);
end
end

function vectorValue = localColumnVector(inputValue, n)
vectorValue = NaN(n, 1);
if isempty(inputValue) || ~isnumeric(inputValue)
    return;
end
raw = double(inputValue(:));
copyCount = min(numel(raw), n);
vectorValue(1:copyCount) = raw(1:copyCount);
end

function feasibleDemand = localClampDemandToAvailable(torqueDemand, posLimit, negLimit)
feasibleDemand = torqueDemand;
driveClamp = isfinite(feasibleDemand) & isfinite(posLimit) & feasibleDemand > posLimit;
regenClamp = isfinite(feasibleDemand) & isfinite(negLimit) & feasibleDemand < negLimit;
feasibleDemand(driveClamp) = posLimit(driveClamp);
feasibleDemand(regenClamp) = negLimit(regenClamp);
end

function propulsionSignal = localPropulsionDeliverySignal(d, tractionForce)
propulsionSignal = tractionForce;
if isstruct(d) && isfield(d, 'finalDriveTorque_Nm')
    candidate = d.finalDriveTorque_Nm(:);
    if any(isfinite(candidate))
        propulsionSignal = candidate;
    end
end
end

function slewRate = localSlewRate(t, signalData)
slewRate = NaN(size(signalData));
if numel(t) < 2 || numel(signalData) < 2
    return;
end
for iSample = 2:numel(signalData)
    if isfinite(signalData(iSample)) && isfinite(signalData(iSample - 1)) && isfinite(t(iSample) - t(iSample - 1)) && abs(t(iSample) - t(iSample - 1)) > eps
        slewRate(iSample) = (signalData(iSample) - signalData(iSample - 1)) / (t(iSample) - t(iSample - 1));
    end
end
end

function ripplePct = localRipplePct(signalData, activeMask)
ripplePct = NaN;
valid = activeMask(:) & isfinite(signalData(:));
if ~any(valid)
    return;
end
sample = signalData(valid);
meanLevel = mean(abs(sample), 'omitnan');
if ~isfinite(meanLevel) || meanLevel <= eps
    return;
end
ripplePct = 100 * std(sample, 'omitnan') / meanLevel;
end

function imbalancePct = localBalanceImbalancePct(signalA, signalB)
imbalancePct = NaN(size(signalA));
valid = isfinite(signalA) & isfinite(signalB) & (abs(signalA) + abs(signalB)) > eps;
imbalancePct(valid) = 100 * abs(signalA(valid) - signalB(valid)) ./ max(abs(signalA(valid)) + abs(signalB(valid)), eps);
end

function dipPct = localShiftTorqueDip(t, torqueSignal, shiftIdx)
dipPct = NaN(numel(shiftIdx), 1);
for iShift = 1:numel(shiftIdx)
    idx = shiftIdx(iShift);
    windowMask = abs(t - t(idx)) <= 1.5;
    preMask = t >= t(idx) - 3.0 & t < t(idx) - 0.25;
    windowTorque = abs(torqueSignal(windowMask));
    preTorque = abs(torqueSignal(preMask));
    if isempty(windowTorque) || isempty(preTorque)
        continue;
    end
    refLevel = median(preTorque, 'omitnan');
    minLevel = min(windowTorque, [], 'omitnan');
    if isfinite(refLevel) && refLevel > eps && isfinite(minLevel)
        dipPct(iShift) = 100 * max(refLevel - minLevel, 0) / refLevel;
    end
end
end

function spike_kW = localShiftLossSpike(t, lossPower, shiftIdx)
spike_kW = NaN(numel(shiftIdx), 1);
for iShift = 1:numel(shiftIdx)
    idx = shiftIdx(iShift);
    windowMask = abs(t - t(idx)) <= 1.5;
    preMask = t >= t(idx) - 3.0 & t < t(idx) - 0.25;
    if ~any(windowMask) || ~any(preMask)
        continue;
    end
    localPeak = max(lossPower(windowMask), [], 'omitnan');
    baseline = median(lossPower(preMask), 'omitnan');
    if isfinite(localPeak) && isfinite(baseline)
        spike_kW(iShift) = localPeak - baseline;
    end
end
end

function plotFile = localPlotDemandCapabilityDelivery(outputFolder, t, torqueDemand, feasibleDemand, actualTorque, posLimit, negLimit, driveShortfall, regenShortfall, propulsionSignal, gear, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(4, 1, 1);
plot(t, torqueDemand, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, feasibleDemand, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', config.Plot.LineWidth);
plot(t, actualTorque, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
ylabel('Torque [Nm]');
title('Demand, Feasible Demand, and Actual Motor Torque');
legend({'Demand', 'Feasible demand', 'Actual', 'Zero'}, 'Location', 'best');
grid on;

subplot(4, 1, 2);
plot(t, posLimit, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, negLimit, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, torqueDemand, ':', 'Color', config.Plot.Colors.Demand, 'LineWidth', 1.0);
ylabel('Torque [Nm]');
title('Available Drive and Regen Torque Envelope');
legend({'Max available', 'Min available', 'Demand'}, 'Location', 'best');
grid on;

subplot(4, 1, 3);
plot(t, driveShortfall, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, regenShortfall, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', config.Plot.Colors.Neutral);
ylabel('Shortfall [Nm]');
title('Drive and Regen Torque Shortfall');
legend({'Drive shortfall', 'Regen shortfall', 'Tracking warning'}, 'Location', 'best');
grid on;

subplot(4, 1, 4);
yyaxis left;
plot(t, propulsionSignal, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
ylabel('Propulsion [Nm or N]');
yyaxis right;
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
ylabel('Gear [-]');
xlabel('Time [s]');
title('Road-Side Delivery and Gear State');
grid on;

sgtitle('Electric Drive Unit Demand-Capability-Delivery Chain');
plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_DemandCapabilityDelivery', config));
close(fig);
end

function plotFile = localPlotEnergyPath(outputFolder, elecDriveEnergy, motorMechDriveEnergy, gbxDriveEnergy, roadDriveEnergy, elecRegenEnergy, roadRegenEnergy, motorLossEnergy, gearboxLossEnergy, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(1, 2, 1);
bar(categorical({'Electrical in', 'Motor shaft', 'Gearbox out', 'Road tractive'}), ...
    [elecDriveEnergy, motorMechDriveEnergy, gbxDriveEnergy, roadDriveEnergy], 'FaceColor', config.Plot.Colors.Vehicle);
ylabel('Energy [kWh]');
title('Drive Energy Conversion Path');
grid on;

subplot(1, 2, 2);
bar(categorical({'Road regen', 'Electrical regen', 'Motor loss', 'Gearbox loss'}), ...
    [roadRegenEnergy, elecRegenEnergy, motorLossEnergy, gearboxLossEnergy], 'FaceColor', config.Plot.Colors.Motor);
ylabel('Energy [kWh]');
title('Regen and Loss Energy');
grid on;

sgtitle('Electric Drive Unit Energy Path');
plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_EnergyPath', config));
close(fig);
end

function plotFile = localPlotTorqueForcePath(outputFolder, t, motorTorque, gbxOutTrq, tractionForce, gear, gearRatio, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, motorTorque, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, gbxOutTrq, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
ylabel('Torque [Nm]');
title('Motor Torque to Gearbox Output Torque');
legend({'Total motor torque', 'Gearbox output torque', 'Zero'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, tractionForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
ylabel('Force [N]');
title('Final Drive Tractive Force Output');
grid on;

subplot(3, 1, 3);
yyaxis left;
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
ylabel('Gear [-]');
yyaxis right;
plot(t, gearRatio, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
ylabel('Gear ratio [-]');
xlabel('Time [s]');
title('Gear State and Ratio Context');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_TorqueForcePath', config));
close(fig);
end

function plotFile = localPlotLossAndShift(outputFolder, t, motorElecPower, motorLossPower, gearboxLossPower, driveLossShare, shiftMask, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, max(motorElecPower, 0), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(-motorElecPower, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
ylabel('Power [kW]');
title('Motor Electrical Drive and Regen Power');
legend({'Drive power', 'Regen power'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, max(motorLossPower, 0), 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(gearboxLossPower, 0), 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
legendText = {'Motor/inverter loss', 'Gearbox loss'};
if any(shiftMask)
    plot(t(shiftMask), max(gearboxLossPower(shiftMask), 0), 'o', 'Color', config.Plot.Colors.Vehicle, 'MarkerSize', 4);
    legendText{end + 1} = 'Shift samples';
end
ylabel('Loss [kW]');
title('Motor/Inverter and Gearbox Loss Power');
legend(legendText, 'Location', 'best');
grid on;

subplot(3, 1, 3);
plot(t, driveLossShare, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.HighLossShare_pct, '--', 'Color', config.Plot.Colors.Neutral);
ylabel('Loss share [%]');
xlabel('Time [s]');
title('EDU Internal Loss Share During Drive');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_LossAndShift', config));
close(fig);
end

function plotFile = localPlotOperatingMap(outputFolder, motorSpeed, motorTorque, gear, driveLossShare, config)
plotFile = "";
valid = isfinite(motorSpeed) & isfinite(motorTorque);
if ~any(valid)
    return;
end
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(1, 2, 1);
gearColor = gear;
gearColor(~isfinite(gearColor)) = 0;
scatter(abs(motorSpeed(valid)), motorTorque(valid), 12, gearColor(valid), 'filled');
xlabel('Motor speed [rpm]');
ylabel('Motor torque [Nm]');
title('Motor Operating Map Colored by Gear');
cb = colorbar;
cb.Label.String = 'Gear';
grid on;

subplot(1, 2, 2);
validLoss = valid & isfinite(driveLossShare);
if any(validLoss)
    scatter(abs(motorSpeed(validLoss)), motorTorque(validLoss), 12, driveLossShare(validLoss), 'filled');
end
xlabel('Motor speed [rpm]');
ylabel('Motor torque [Nm]');
title('Operating Map Colored by EDU Loss Share');
cb = colorbar;
cb.Label.String = 'Loss share [%]';
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_OperatingMap', config));
close(fig);
end

function plotFile = localPlotMotorBalance(outputFolder, t, em1Trq, em2Trq, em1Loss, em2Loss, speedMismatchRpm, torqueImbalancePct, config)
plotFile = "";
hasTorque = any(isfinite(em1Trq)) && any(isfinite(em2Trq));
hasLoss = any(isfinite(em1Loss)) && any(isfinite(em2Loss));
hasSpeed = any(isfinite(speedMismatchRpm));
if ~hasTorque && ~hasLoss && ~hasSpeed
    return;
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
if hasTorque
    scatter(em1Trq, em2Trq, 12, config.Plot.Colors.Motor, 'filled'); hold on;
    lim = max(abs([em1Trq(:); em2Trq(:)]), [], 'omitnan');
    if isfinite(lim) && lim > 0
        plot([-lim lim], [-lim lim], '--', 'Color', config.Plot.Colors.Neutral);
        xlim([-lim lim]);
        ylim([-lim lim]);
    end
end
xlabel('Motor 1 torque [Nm]');
ylabel('Motor 2 torque [Nm]');
title('Motor Torque Sharing');
grid on;

subplot(2, 2, 2);
plot(t, torqueImbalancePct, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
ylabel('Imbalance [%]');
title('Torque Sharing Imbalance');
grid on;

subplot(2, 2, 3);
if hasLoss
    plot(t, em1Loss, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
    plot(t, em2Loss, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    legend({'Motor 1 loss', 'Motor 2 loss'}, 'Location', 'best');
end
ylabel('Loss [kW]');
xlabel('Time [s]');
title('Motor Loss Balance');
grid on;

subplot(2, 2, 4);
plot(t, speedMismatchRpm, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
ylabel('Speed mismatch [rpm]');
xlabel('Time [s]');
title('Motor Speed Consistency');
grid on;

sgtitle('Electric Drive Unit Dual-Motor Balance');
plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_MotorBalance', config));
close(fig);
end

function plotFile = localPlotShiftDisturbance(outputFolder, t, shiftIdx, shiftTorqueDipPct, shiftLossSpike_kW, gbxOutTrq, propulsionSignal, gear, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
if isempty(shiftIdx)
    text(0.5, 0.5, 'No shift events detected', 'HorizontalAlignment', 'center');
    axis off;
else
    bar(1:numel(shiftIdx), shiftTorqueDipPct, 'FaceColor', config.Plot.Colors.Vehicle);
    ylabel('Dip [%]');
    title('Torque Dip by Shift Event');
    grid on;
end

subplot(3, 1, 2);
if isempty(shiftIdx)
    text(0.5, 0.5, 'No shift events detected', 'HorizontalAlignment', 'center');
    axis off;
else
    bar(1:numel(shiftIdx), shiftLossSpike_kW, 'FaceColor', config.Plot.Colors.Warning);
    ylabel('Loss spike [kW]');
    title('Loss Spike by Shift Event');
    grid on;
end

subplot(3, 1, 3);
plot(t, gbxOutTrq, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, propulsionSignal, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
for iShift = 1:numel(shiftIdx)
    xline(t(shiftIdx(iShift)), ':', 'Color', config.Plot.Colors.Gear);
end
yyaxis right;
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', 0.9);
ylabel('Gear [-]');
yyaxis left;
ylabel('Torque / force');
xlabel('Time [s]');
title('Shift Context on Gearbox Output and Road-Side Delivery');
legend({'Gearbox output torque', 'Road-side delivery'}, 'Location', 'best');
grid on;

sgtitle('Electric Drive Unit Shift Disturbance');
plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_ShiftDisturbance', config));
close(fig);
end

function plotFiles = localAppendPlotFile(plotFiles, plotFile)
if nargin < 2 || strlength(string(plotFile)) == 0
    return;
end
plotFiles(end + 1, 1) = string(plotFile);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end

function subResult = localRunLegacySubmodule(functionHandle, displayName, analysisData, outputPaths, config)
try
    subResult = functionHandle(analysisData, outputPaths, config);
catch subException
    subResult = localInitResult(displayName, {}, {});
    subResult.Warnings = "Legacy " + displayName + " sub-analysis was skipped inside Electric Drive Unit: " + string(subException.message);
end
end
