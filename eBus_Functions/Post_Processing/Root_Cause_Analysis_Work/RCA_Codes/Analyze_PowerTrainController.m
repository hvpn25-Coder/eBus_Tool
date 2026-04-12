function result = Analyze_PowerTrainController(analysisData, outputPaths, config)
% Analyze_PowerTrainController  Pedal-to-torque command quality and saturation analysis.

result = localInitResult("POWER TRAIN CONTROLLER", ...
    {'acc_pdl', 'brk_pdl', 'emot1_dem_trq', 'emot2_dem_trq'}, ...
    {'emot1_act_trq', 'emot2_act_trq', 'max_emot1_dem_trq', 'max_emot2_dem_trq', ...
    'min_emot1_dem_trq', 'min_emot2_dem_trq', 'gr_num', 'gr_ratio', ...
    'batt_chrg_pwr_lim', 'batt_dischrg_pwr_lim', 'batt_soc', 'batt_pwr', ...
    'emot1_pwr', 'emot2_pwr', 'emot1_act_spd', 'emot2_act_spd'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Powertrain controller analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.SummaryText = summary;
    result.Suggestions = RCA_MakeSuggestionTable("Power Train Controller", recs, evidence);
    return;
end

accPedal = d.accPedal_pct(:);
brkPedal = d.brkPedal_pct(:);
dt = d.dt_s(:);
vehSpeed = d.vehVel_kmh(:);
totalDemand = d.torqueDemandTotal_Nm(:);
totalActual = d.torqueActualTotal_Nm(:);
posLimit = d.controllerTorquePositiveLimit_Nm(:);
negLimit = d.controllerTorqueNegativeLimit_Nm(:);
gearNum = d.gearNumber(:);
gearRatio = d.gearRatio(:);
battSoc = d.batterySOC_pct(:);
battDischargeLimit = abs(d.battDischargePowerLimit_kW(:));
battChargeLimit = abs(d.battChargePowerLimit_kW(:));
motorElecPwr = d.motorElectricalPower_kW(:);
motorDriveElecPwr = d.motorDriveElectricalPowerPositive_kW(:);
motorRegenElecPwr = d.motorRegenElectricalPowerPositive_kW(:);

if ~any(isfinite(gearNum))
    gearNum = localAlignedSignal(analysisData.Signals, 'gr_num', n);
end
if ~any(isfinite(gearRatio))
    gearRatio = localAlignedSignal(analysisData.Signals, 'gr_ratio', n);
end

em1Demand = localAlignedSignal(analysisData.Signals, 'emot1_dem_trq', n);
em2Demand = localAlignedSignal(analysisData.Signals, 'emot2_dem_trq', n);
em1Actual = localAlignedSignal(analysisData.Signals, 'emot1_act_trq', n);
em2Actual = localAlignedSignal(analysisData.Signals, 'emot2_act_trq', n);
em1Max = localAlignedSignal(analysisData.Signals, 'max_emot1_dem_trq', n);
em2Max = localAlignedSignal(analysisData.Signals, 'max_emot2_dem_trq', n);
em1Min = localAlignedSignal(analysisData.Signals, 'min_emot1_dem_trq', n);
em2Min = localAlignedSignal(analysisData.Signals, 'min_emot2_dem_trq', n);

activeTorqueThreshold = config.Thresholds.ControllerTorqueDeadband_Nm;
leakageTorqueThreshold = config.Thresholds.ControllerLeakageTorque_Nm;
torqueRate = localSignalRate(totalDemand, dt);
torqueError = totalDemand - totalActual;

validDemand = isfinite(totalDemand);
movingMask = isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh;
accelPhase = isfinite(accPedal) & accPedal >= config.Thresholds.DriverPedalActive_pct & ...
    ~(isfinite(brkPedal) & brkPedal >= config.Thresholds.DriverPedalActive_pct);
brakePhase = isfinite(brkPedal) & brkPedal >= config.Thresholds.DriverPedalActive_pct;
cruisePhase = movingMask & ~accelPhase & ~brakePhase;

driveMask = validDemand & totalDemand > activeTorqueThreshold;
regenMask = validDemand & totalDemand < -activeTorqueThreshold;
neutralMask = validDemand & abs(totalDemand) <= activeTorqueThreshold;

validPosLimit = driveMask & isfinite(posLimit) & posLimit > 0;
validNegLimit = regenMask & isfinite(negLimit) & abs(negLimit) > 0;
driveNearLimitMask = validPosLimit & totalDemand >= config.Thresholds.LimitUsageFraction .* posLimit;
regenNearLimitMask = validNegLimit & abs(totalDemand) >= config.Thresholds.LimitUsageFraction .* abs(negLimit);
driveReserve = posLimit - totalDemand;
regenReserve = abs(negLimit) - abs(totalDemand);
driveLimitViolation = validPosLimit & totalDemand > posLimit + activeTorqueThreshold;
regenLimitViolation = validNegLimit & abs(totalDemand) > abs(negLimit) + activeTorqueThreshold;

driveShortfall = NaN(size(totalDemand));
regenShortfall = NaN(size(totalDemand));
validActual = isfinite(totalDemand) & isfinite(totalActual);
driveShortfall(validActual & driveMask) = max(totalDemand(validActual & driveMask) - totalActual(validActual & driveMask), 0);
regenShortfall(validActual & regenMask) = max(totalActual(validActual & regenMask) - totalDemand(validActual & regenMask), 0);
signMismatchMask = validActual & abs(totalDemand) > activeTorqueThreshold & abs(totalActual) > activeTorqueThreshold & sign(totalDemand) ~= sign(totalActual);

validDischargePowerLimit = isfinite(motorDriveElecPwr) & isfinite(battDischargeLimit) & battDischargeLimit > 0 & motorDriveElecPwr > 0;
validChargePowerLimit = isfinite(motorRegenElecPwr) & isfinite(battChargeLimit) & battChargeLimit > 0 & motorRegenElecPwr > 0;
drivePowerUse = NaN(size(t));
regenPowerUse = NaN(size(t));
drivePowerUse(validDischargePowerLimit) = motorDriveElecPwr(validDischargePowerLimit) ./ max(battDischargeLimit(validDischargePowerLimit), eps);
regenPowerUse(validChargePowerLimit) = motorRegenElecPwr(validChargePowerLimit) ./ max(battChargeLimit(validChargePowerLimit), eps);
nearDischargePowerLimit = validDischargePowerLimit & drivePowerUse >= config.Thresholds.LimitUsageFraction;
nearChargePowerLimit = validChargePowerLimit & regenPowerUse >= config.Thresholds.LimitUsageFraction;

brakePhasePositiveTorque = brakePhase & totalDemand > activeTorqueThreshold;
accelPhaseRegenTorque = accelPhase & totalDemand < -activeTorqueThreshold;
cruiseLeakageMask = cruisePhase & abs(totalDemand) > leakageTorqueThreshold;

splitImbalanceDrive = localSplitImbalance(em1Demand, em2Demand, totalDemand, driveMask);
splitImbalanceRegen = localSplitImbalance(em1Demand, em2Demand, totalDemand, regenMask);
em1DriveNearLimit = localMotorNearLimit(em1Demand, em1Max, driveMask, config.Thresholds.LimitUsageFraction, true);
em2DriveNearLimit = localMotorNearLimit(em2Demand, em2Max, driveMask, config.Thresholds.LimitUsageFraction, true);
em1RegenNearLimit = localMotorNearLimit(em1Demand, em1Min, regenMask, config.Thresholds.LimitUsageFraction, false);
em2RegenNearLimit = localMotorNearLimit(em2Demand, em2Min, regenMask, config.Thresholds.LimitUsageFraction, false);
validGear = isfinite(gearNum);
shiftMask = localShiftMask(gearNum);
gearShiftCount = sum(localShiftEvents(gearNum));

[accelTorqueGain, accelTorqueCorr, accelNonlinearity, accelDeadband] = localPedalTorqueMappingStats(accPedal, max(totalDemand, 0), accelPhase, activeTorqueThreshold);
[brakeRegenGain, brakeRegenCorr, brakeNonlinearity, brakeDeadband] = localPedalTorqueMappingStats(brkPedal, max(-totalDemand, 0), brakePhase, activeTorqueThreshold);
commandReversalCount = localCommandReversalCount(totalDemand, activeTorqueThreshold);

rows = RCA_AddKPI(rows, 'Drive Command Active Share', 100 * RCA_FractionTrue(driveMask, validDemand), '%', ...
    'Operation', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    sprintf('Total demand above %.1f Nm is treated as an active drive request.', activeTorqueThreshold));
rows = RCA_AddKPI(rows, 'Regen Command Active Share', 100 * RCA_FractionTrue(regenMask, validDemand), '%', ...
    'Operation', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    sprintf('Total demand below -%.1f Nm is treated as an active recuperation request.', activeTorqueThreshold));
rows = RCA_AddKPI(rows, 'Neutral Command Share', 100 * RCA_FractionTrue(neutralMask, validDemand), '%', ...
    'Operation', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    'Near-zero torque request share shows how often the controller commands a neutral or coast state.');
rows = RCA_AddKPI(rows, 'Peak Drive Torque Demand', max(totalDemand, [], 'omitnan'), 'Nm', ...
    'Performance', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    'Peak positive total torque request from the controller.');
rows = RCA_AddKPI(rows, 'Peak Regen Torque Demand', max(-totalDemand, [], 'omitnan'), 'Nm', ...
    'Efficiency', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    'Peak recuperation torque magnitude requested by the controller.');
rows = RCA_AddKPI(rows, 'Accelerator-to-Drive Torque Gain', accelTorqueGain, 'Nm/%', ...
    'Demand Mapping', 'Power Train Controller', 'acc_pdl + total demand torque', ...
    'Linearized gain from accelerator pedal to positive torque demand over active acceleration samples.');
rows = RCA_AddKPI(rows, 'Accelerator-to-Drive Correlation', accelTorqueCorr, '-', ...
    'Demand Mapping', 'Power Train Controller', 'acc_pdl + total demand torque', ...
    'Manual Pearson correlation; values near 1 indicate consistent monotonic pedal-to-drive torque mapping.');
rows = RCA_AddKPI(rows, 'Accelerator Mapping Nonlinearity Index', accelNonlinearity, '-', ...
    'Demand Mapping', 'Power Train Controller', 'acc_pdl + total demand torque', ...
    'Normalized residual of linear pedal-to-torque fit; high values indicate progressive mapping, clipping, or discontinuity.');
rows = RCA_AddKPI(rows, 'Accelerator Deadband Estimate', accelDeadband, '% pedal', ...
    'Demand Mapping', 'Power Train Controller', 'acc_pdl + total demand torque', ...
    sprintf('Estimated from the lowest active-pedal region where drive torque exceeds %.1f Nm.', activeTorqueThreshold));
rows = RCA_AddKPI(rows, 'Brake-to-Regen Torque Gain', brakeRegenGain, 'Nm/%', ...
    'Demand Mapping', 'Power Train Controller', 'brk_pdl + total demand torque', ...
    'Linearized gain from brake pedal to recuperation torque magnitude over active braking samples.');
rows = RCA_AddKPI(rows, 'Brake-to-Regen Correlation', brakeRegenCorr, '-', ...
    'Demand Mapping', 'Power Train Controller', 'brk_pdl + total demand torque', ...
    'Manual Pearson correlation; values near 1 indicate consistent brake-to-regen mapping.');
rows = RCA_AddKPI(rows, 'Brake Regen Mapping Nonlinearity Index', brakeNonlinearity, '-', ...
    'Demand Mapping', 'Power Train Controller', 'brk_pdl + total demand torque', ...
    'Normalized residual of brake-to-regen linear fit; high values indicate blending, limit clipping, or calibration discontinuity.');
rows = RCA_AddKPI(rows, 'Brake Regen Deadband Estimate', brakeDeadband, '% pedal', ...
    'Demand Mapping', 'Power Train Controller', 'brk_pdl + total demand torque', ...
    sprintf('Estimated from the lowest active brake-pedal region where regen torque exceeds %.1f Nm.', activeTorqueThreshold));
rows = RCA_AddKPI(rows, 'Torque Command Rate RMS', localRmsFinite(torqueRate, validDemand), 'Nm/s', ...
    'Drivability', 'Power Train Controller', 'torque demand + time', ...
    'RMS torque command rate is a drivability proxy for torque smoothness and tip-in/tip-out aggressiveness.');
rows = RCA_AddKPI(rows, 'Torque Command Rate 95th Percentile', RCA_Percentile(abs(torqueRate(validDemand & isfinite(torqueRate))), 95), 'Nm/s', ...
    'Drivability', 'Power Train Controller', 'torque demand + time', ...
    'Tail command-rate metric highlights abrupt torque steps, clipping, or mode switching.');
rows = RCA_AddKPI(rows, 'Torque Command Reversal Count', commandReversalCount, 'count', ...
    'Drivability', 'Power Train Controller', 'torque demand', ...
    'Counts drive-to-regen or regen-to-drive command sign changes outside the deadband.');

if any(validPosLimit)
    rows = RCA_AddKPI(rows, 'Near Positive Limit Share', 100 * RCA_FractionTrue(driveNearLimitMask, validPosLimit), '%', ...
        'Saturation', 'Power Train Controller', 'demand torque + max available torque', ...
        sprintf('Demand above %.0f%% of the positive limit is treated as near-limit drive operation.', config.Thresholds.LimitUsageFraction * 100));
    rows = RCA_AddKPI(rows, 'Mean Positive Torque Reserve', mean(driveReserve(validPosLimit), 'omitnan'), 'Nm', ...
        'Saturation', 'Power Train Controller', 'demand torque + max available torque', ...
        'Positive torque reserve is max available drive torque minus demanded drive torque.');
    rows = RCA_AddKPI(rows, 'Positive Torque Limit Violation Share', 100 * RCA_FractionTrue(driveLimitViolation, validPosLimit), '%', ...
        'Saturation', 'Power Train Controller', 'demand torque + max available torque', ...
        'Demand above the positive envelope indicates clipping risk, stale limits, or sign/mapping issue.');
end

if any(validNegLimit)
    rows = RCA_AddKPI(rows, 'Near Regen Limit Share', 100 * RCA_FractionTrue(regenNearLimitMask, validNegLimit), '%', ...
        'Saturation', 'Power Train Controller', 'demand torque + min available torque', ...
        sprintf('Demand above %.0f%% of the recuperation limit magnitude is treated as near-limit regen operation.', config.Thresholds.LimitUsageFraction * 100));
    rows = RCA_AddKPI(rows, 'Mean Regen Torque Reserve', mean(regenReserve(validNegLimit), 'omitnan'), 'Nm', ...
        'Saturation', 'Power Train Controller', 'demand torque + min available torque', ...
        'Regen reserve is available recuperation torque magnitude minus demanded recuperation torque magnitude.');
    rows = RCA_AddKPI(rows, 'Regen Torque Limit Violation Share', 100 * RCA_FractionTrue(regenLimitViolation, validNegLimit), '%', ...
        'Saturation', 'Power Train Controller', 'demand torque + min available torque', ...
        'Demand beyond the negative envelope indicates regen clipping risk, stale limits, or sign/mapping issue.');
end

if any(validActual)
    rows = RCA_AddKPI(rows, 'Total Torque Tracking MAE', mean(abs(totalDemand(validActual) - totalActual(validActual)), 'omitnan'), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Mean absolute total torque tracking error between commanded and delivered torque.');
    rows = RCA_AddKPI(rows, 'Total Torque Tracking RMSE', sqrt(mean(torqueError(validActual).^2, 'omitnan')), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Root-mean-square total torque tracking error; emphasizes large transient misses.');
    rows = RCA_AddKPI(rows, 'Peak Absolute Torque Tracking Error', max(abs(torqueError(validActual)), [], 'omitnan'), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Worst instantaneous difference between demanded and actual total motor torque.');
    rows = RCA_AddKPI(rows, 'Drive Torque Shortfall 95th Percentile', RCA_Percentile(driveShortfall(isfinite(driveShortfall)), 95), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Positive torque under-delivery tail severity during propulsion.');
    rows = RCA_AddKPI(rows, 'Regen Torque Shortfall 95th Percentile', RCA_Percentile(regenShortfall(isfinite(regenShortfall)), 95), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Recuperation torque under-delivery tail severity during braking or lift-off regen.');
    rows = RCA_AddKPI(rows, 'Torque Sign Mismatch Share', 100 * RCA_FractionTrue(signMismatchMask, validActual), '%', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Demand and actual torque have opposite signs outside the torque deadband; useful for detecting handover or sign-convention issues.');
    rows = RCA_AddKPI(rows, 'Motor 1 Torque Tracking MAE', mean(abs(em1Demand - em1Actual), 'omitnan'), 'Nm', ...
        'Tracking', 'Power Train Controller', 'emot1_dem_trq + emot1_act_trq', ...
        'Motor 1 demand-to-actual torque tracking quality.');
    rows = RCA_AddKPI(rows, 'Motor 2 Torque Tracking MAE', mean(abs(em2Demand - em2Actual), 'omitnan'), 'Nm', ...
        'Tracking', 'Power Train Controller', 'emot2_dem_trq + emot2_act_trq', ...
        'Motor 2 demand-to-actual torque tracking quality.');
end

if any(validDischargePowerLimit)
    rows = RCA_AddKPI(rows, 'Drive Power Utilization Mean', mean(drivePowerUse(validDischargePowerLimit), 'omitnan') * 100, '%', ...
        'Power Utilization', 'Power Train Controller', 'emot power + batt_dischrg_pwr_lim', ...
        'Mean actual electric drive power as a share of BMS discharge power limit.');
    rows = RCA_AddKPI(rows, 'Drive Power Utilization 95th Percentile', RCA_Percentile(drivePowerUse(validDischargePowerLimit) * 100, 95), '%', ...
        'Power Utilization', 'Power Train Controller', 'emot power + batt_dischrg_pwr_lim', ...
        'Tail usage of available discharge power; high values indicate battery-power-limited acceleration risk.');
    rows = RCA_AddKPI(rows, 'Near BMS Discharge Power Limit Share', 100 * RCA_FractionTrue(nearDischargePowerLimit, validDischargePowerLimit), '%', ...
        'Power Utilization', 'Power Train Controller', 'emot power + batt_dischrg_pwr_lim', ...
        sprintf('Drive power above %.0f%% of BMS discharge limit is treated as near battery power limit.', config.Thresholds.LimitUsageFraction * 100));
end

if any(validChargePowerLimit)
    rows = RCA_AddKPI(rows, 'Regen Power Utilization Mean', mean(regenPowerUse(validChargePowerLimit), 'omitnan') * 100, '%', ...
        'Power Utilization', 'Power Train Controller', 'emot power + batt_chrg_pwr_lim', ...
        'Mean actual electric recuperation power as a share of BMS charge power limit.');
    rows = RCA_AddKPI(rows, 'Regen Power Utilization 95th Percentile', RCA_Percentile(regenPowerUse(validChargePowerLimit) * 100, 95), '%', ...
        'Power Utilization', 'Power Train Controller', 'emot power + batt_chrg_pwr_lim', ...
        'Tail usage of available charge power; high values indicate regen curtailment risk.');
    rows = RCA_AddKPI(rows, 'Near BMS Charge Power Limit Share', 100 * RCA_FractionTrue(nearChargePowerLimit, validChargePowerLimit), '%', ...
        'Power Utilization', 'Power Train Controller', 'emot power + batt_chrg_pwr_lim', ...
        sprintf('Regen power above %.0f%% of BMS charge limit is treated as near battery charge limit.', config.Thresholds.LimitUsageFraction * 100));
end

if any(isfinite(motorElecPwr))
    rows = RCA_AddKPI(rows, 'Motor Electrical Drive Energy', RCA_TrapzFinite(t, max(motorElecPwr, 0)) / 3600, 'kWh', ...
        'Energy', 'Power Train Controller', 'emot1_pwr + emot2_pwr', ...
        'Integrated positive motor electrical power requested/delivered through the eDrive system.');
    rows = RCA_AddKPI(rows, 'Motor Electrical Regen Energy', RCA_TrapzFinite(t, max(-motorElecPwr, 0)) / 3600, 'kWh', ...
        'Energy', 'Power Train Controller', 'emot1_pwr + emot2_pwr', ...
        'Integrated negative motor electrical power magnitude during recuperation.');
end

rows = RCA_AddKPI(rows, 'Positive Torque During Brake Phase Share', 100 * RCA_FractionTrue(brakePhasePositiveTorque, brakePhase), '%', ...
    'Arbitration', 'Power Train Controller', 'brk_pdl + total demand torque', ...
    'Positive drive torque while braking is active is a sign-conflict or arbitration-quality indicator.');
rows = RCA_AddKPI(rows, 'Regen Torque During Accel Phase Share', 100 * RCA_FractionTrue(accelPhaseRegenTorque, accelPhase), '%', ...
    'Arbitration', 'Power Train Controller', 'acc_pdl + total demand torque', ...
    'Recuperation torque while acceleration is active is a sign-conflict or handover-quality indicator.');
rows = RCA_AddKPI(rows, 'Cruise Torque Leakage Share', 100 * RCA_FractionTrue(cruiseLeakageMask, cruisePhase), '%', ...
    'Operation', 'Power Train Controller', 'pedals + total demand torque', ...
    sprintf('Residual command above %.1f Nm during nominal cruise/coast is treated as unwanted torque leakage.', leakageTorqueThreshold));

rows = RCA_AddKPI(rows, 'Mean Drive Split Imbalance', mean(splitImbalanceDrive, 'omitnan'), '%', ...
    'Allocation', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    sprintf('Motor demand imbalance above %.1f%% is worth reviewing for split logic or actuator capability asymmetry.', config.Thresholds.ControllerSplitImbalance_pct));
rows = RCA_AddKPI(rows, 'Mean Regen Split Imbalance', mean(splitImbalanceRegen, 'omitnan'), '%', ...
    'Allocation', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', ...
    'Motor demand imbalance during recuperation shows whether regen is distributed evenly or intentionally biased.');
rows = RCA_AddKPI(rows, 'Motor 1 Near Drive Limit Share', em1DriveNearLimit, '%', ...
    'Allocation', 'Power Train Controller', 'emot1_dem_trq + max_emot1_dem_trq', ...
    'Share of active drive samples where Motor 1 is commanded near its positive torque limit.');
rows = RCA_AddKPI(rows, 'Motor 2 Near Drive Limit Share', em2DriveNearLimit, '%', ...
    'Allocation', 'Power Train Controller', 'emot2_dem_trq + max_emot2_dem_trq', ...
    'Share of active drive samples where Motor 2 is commanded near its positive torque limit.');
rows = RCA_AddKPI(rows, 'Motor 1 Near Regen Limit Share', em1RegenNearLimit, '%', ...
    'Allocation', 'Power Train Controller', 'emot1_dem_trq + min_emot1_dem_trq', ...
    'Share of active regen samples where Motor 1 is commanded near its recuperation limit.');
rows = RCA_AddKPI(rows, 'Motor 2 Near Regen Limit Share', em2RegenNearLimit, '%', ...
    'Allocation', 'Power Train Controller', 'emot2_dem_trq + min_emot2_dem_trq', ...
    'Share of active regen samples where Motor 2 is commanded near its recuperation limit.');

rows = RCA_AddKPI(rows, 'Mean Acceleration-Phase Torque Demand', mean(totalDemand(accelPhase), 'omitnan'), 'Nm', ...
    'Phases', 'Power Train Controller', 'acc_pdl + total demand torque', ...
    'Average torque request while accelerator pedal is active.');
rows = RCA_AddKPI(rows, 'Mean Braking-Phase Regen Demand', mean(max(-totalDemand(brakePhase), 0), 'omitnan'), 'Nm', ...
    'Phases', 'Power Train Controller', 'brk_pdl + total demand torque', ...
    'Average recuperation torque magnitude requested while brake pedal is active.');
rows = RCA_AddKPI(rows, 'Mean Cruise-Phase Absolute Torque', mean(abs(totalDemand(cruisePhase)), 'omitnan'), 'Nm', ...
    'Phases', 'Power Train Controller', 'pedals + total demand torque', ...
    'Average absolute torque during nominal cruise/coast helps identify unwanted residual request.');

if any(validGear)
    rows = RCA_AddKPI(rows, 'Shift Count', gearShiftCount, 'count', ...
        'Gear Context', 'Power Train Controller', 'gr_num', ...
        'Total number of actual gear changes observed during the analysis window.');
    rows = RCA_AddKPI(rows, 'Shift Active Share', 100 * RCA_FractionTrue(shiftMask, validGear), '%', ...
        'Gear Context', 'Power Train Controller', 'gr_num', ...
        'Share of samples around a gear transition. High values indicate frequent gear state changes in controller operating context.');
    rows = RCA_AddKPI(rows, 'Mean Active Gear Ratio', mean(gearRatio(validGear & (driveMask | regenMask)), 'omitnan'), '-', ...
        'Gear Context', 'Power Train Controller', 'gr_ratio', ...
        'Average actual gear ratio while non-neutral torque demand is active.');
    rows = RCA_AddKPI(rows, 'Gear Ratio Span', max(gearRatio(validGear), [], 'omitnan') - min(gearRatio(validGear), [], 'omitnan'), '-', ...
        'Gear Context', 'Power Train Controller', 'gr_ratio', ...
        'Range between maximum and minimum observed actual gear ratio.');
    rows = RCA_AddKPI(rows, 'Near Positive Limit During Shift Share', 100 * RCA_FractionTrue(driveNearLimitMask & shiftMask, shiftMask & driveMask), '%', ...
        'Gear Context', 'Power Train Controller', 'gr_num + max available torque', ...
        'Shows whether drive torque requests are concentrated near machine limit during shift activity.');
    rows = RCA_AddKPI(rows, 'Near Regen Limit During Shift Share', 100 * RCA_FractionTrue(regenNearLimitMask & shiftMask, shiftMask & regenMask), '%', ...
        'Gear Context', 'Power Train Controller', 'gr_num + min available torque', ...
        'Shows whether recuperation torque requests are concentrated near machine limit during shift activity.');

    [gearList, gearDriveMean, gearRegenMean, gearDriveShare] = localGearDemandStats(gearNum, totalDemand, driveMask, regenMask);
    for iGear = 1:numel(gearList)
        gearLabel = sprintf('Gear %.0f', gearList(iGear));
        rows = RCA_AddKPI(rows, [gearLabel ' Drive Command Share'], gearDriveShare(iGear), '%', ...
            'Gear Context', 'Power Train Controller', 'gr_num + total demand torque', ...
            'Share of active drive-demand samples occurring in this actual gear.');
        rows = RCA_AddKPI(rows, [gearLabel ' Mean Drive Torque Demand'], gearDriveMean(iGear), 'Nm', ...
            'Gear Context', 'Power Train Controller', 'gr_num + total demand torque', ...
            'Average total positive torque demand while this gear is active.');
        rows = RCA_AddKPI(rows, [gearLabel ' Mean Regen Torque Demand'], gearRegenMean(iGear), 'Nm', ...
            'Gear Context', 'Power Train Controller', 'gr_num + total demand torque', ...
            'Average recuperation torque magnitude while this gear is active.');
    end
end

summary(end + 1) = sprintf(['Powertrain controller command mix: drive torque is active for %.1f%% of valid samples, ', ...
    'regen torque for %.1f%%, and neutral/coast torque for %.1f%%.'], ...
    100 * RCA_FractionTrue(driveMask, validDemand), 100 * RCA_FractionTrue(regenMask, validDemand), 100 * RCA_FractionTrue(neutralMask, validDemand));

summary(end + 1) = sprintf(['Pedal interpretation quality: accelerator-to-drive gain is %.1f Nm/%% with correlation %.2f; ', ...
    'brake-to-regen gain is %.1f Nm/%% with correlation %.2f. These values describe how the controller converts driver intent into torque demand.'], ...
    accelTorqueGain, accelTorqueCorr, brakeRegenGain, brakeRegenCorr);

if any(validDischargePowerLimit) || any(validChargePowerLimit)
    summary(end + 1) = sprintf(['Power utilization context: mean BMS discharge-limit utilization is %.1f%% and mean charge-limit utilization is %.1f%%. ', ...
        'This separates controller mapping issues from battery power-limit constraints.'], ...
        mean(drivePowerUse(validDischargePowerLimit), 'omitnan') * 100, mean(regenPowerUse(validChargePowerLimit), 'omitnan') * 100);
end

if any(validPosLimit) || any(validNegLimit)
    summary(end + 1) = sprintf(['Controller saturation context: near positive limit share is %.1f%% and near regen limit share is %.1f%%. ', ...
        'High values mean the controller is frequently asking for the edge of available machine capability.'], ...
        100 * RCA_FractionTrue(driveNearLimitMask, validPosLimit), 100 * RCA_FractionTrue(regenNearLimitMask, validNegLimit));
end

summary(end + 1) = sprintf(['Arbitration quality: positive torque during brake phase is %.1f%% and regen torque during accel phase is %.1f%%. ', ...
    'These are useful stakeholder-facing indicators for pedal-to-torque sign consistency.'], ...
    100 * RCA_FractionTrue(brakePhasePositiveTorque, brakePhase), 100 * RCA_FractionTrue(accelPhaseRegenTorque, accelPhase));

if isfinite(mean(splitImbalanceDrive, 'omitnan')) || isfinite(mean(splitImbalanceRegen, 'omitnan'))
    summary(end + 1) = sprintf(['Motor split quality: mean drive split imbalance is %.1f%% and mean regen split imbalance is %.1f%%. ', ...
        'Large imbalance can be intentional, but it should correlate with machine limits, efficiency strategy, or drivetrain architecture.'], ...
        mean(splitImbalanceDrive, 'omitnan'), mean(splitImbalanceRegen, 'omitnan'));
end

if any(validGear)
    [dominantGear, dominantGearShare] = localDominantGear(gearNum, driveMask | regenMask);
    summary(end + 1) = sprintf(['Gear context: %d gear changes were observed, with %.1f%% of active torque-command time spent in Gear %.0f. ', ...
        'This helps stakeholders connect controller torque requests to the actual transmission state used by the feedforward logic.'], ...
        gearShiftCount, dominantGearShare, dominantGear);
    if any(isfinite(gearRatio))
        summary(end + 1) = sprintf(['Actual gear ratio ranged from %.2f to %.2f, with a mean active ratio of %.2f. ', ...
            'This indicates the spread of transmission leverage seen by the controller while generating drive and regen requests.'], ...
            min(gearRatio(validGear), [], 'omitnan'), max(gearRatio(validGear), [], 'omitnan'), ...
            mean(gearRatio(validGear & (driveMask | regenMask)), 'omitnan'));
    end
end

recs(end + 1) = "Acceleration-phase tuning hint: tune pedal-to-drive torque gain, torque ramp rate, and torque-limit approach together so the controller requests strong propulsion without chattering at the positive limit or creating a motor split bias.";
evidence(end + 1) = "General acceleration guidance from accel_pdl, drive-demand share, positive-limit usage, and motor split imbalance.";
recs(end + 1) = "Braking-phase tuning hint: tune brake-pedal to recuperation mapping together with the available negative torque envelope so braking demand is converted into stable regen torque before friction braking takes over.";
evidence(end + 1) = "General braking guidance from brk_pdl, regen-demand share, regen-limit usage, and braking-phase torque sign consistency.";
recs(end + 1) = "Cruise-phase tuning hint: target low residual torque, low sign conflict, and smooth zero-crossing behaviour so the controller does not inject unnecessary drive or regen torque during nominal coasting or steady-state running.";
evidence(end + 1) = "General cruise guidance from neutral-command share, cruise torque leakage share, and phase torque averages.";

if 100 * RCA_FractionTrue(driveNearLimitMask, validPosLimit) > 20
    recs(end + 1) = "Separate true torque-capability limitation from controller over-requesting; frequent positive-limit operation means poor acceleration is not only a PI or pedal-gain issue.";
    evidence(end + 1) = sprintf('Near positive limit share is %.1f%%.', 100 * RCA_FractionTrue(driveNearLimitMask, validPosLimit));
end
if 100 * RCA_FractionTrue(regenNearLimitMask, validNegLimit) > 20
    recs(end + 1) = "Review brake-pedal to regen scaling and negative torque reserve management; the controller frequently requests near-maximum recuperation torque.";
    evidence(end + 1) = sprintf('Near regen limit share is %.1f%%.', 100 * RCA_FractionTrue(regenNearLimitMask, validNegLimit));
end
if 100 * RCA_FractionTrue(brakePhasePositiveTorque, brakePhase) > 2 || 100 * RCA_FractionTrue(accelPhaseRegenTorque, accelPhase) > 2
    recs(end + 1) = "Tighten pedal arbitration and torque sign handover so positive torque is not sustained into braking and regen torque is not sustained into acceleration demand.";
    evidence(end + 1) = sprintf('Brake-phase positive torque share is %.1f%% and accel-phase regen share is %.1f%%.', ...
        100 * RCA_FractionTrue(brakePhasePositiveTorque, brakePhase), 100 * RCA_FractionTrue(accelPhaseRegenTorque, accelPhase));
end
if mean(splitImbalanceDrive, 'omitnan') > config.Thresholds.ControllerSplitImbalance_pct || ...
        mean(splitImbalanceRegen, 'omitnan') > config.Thresholds.ControllerSplitImbalance_pct
    recs(end + 1) = "Review Motor 1 / Motor 2 torque split logic so sustained command imbalance is traceable to an intended capability, protection, or efficiency strategy rather than hidden allocation bias.";
    evidence(end + 1) = sprintf('Mean drive split imbalance is %.1f%% and mean regen split imbalance is %.1f%%.', ...
        mean(splitImbalanceDrive, 'omitnan'), mean(splitImbalanceRegen, 'omitnan'));
end
if 100 * RCA_FractionTrue(cruiseLeakageMask, cruisePhase) > 10
    recs(end + 1) = "Reduce residual torque leakage in cruise/coast by tightening zero-torque calibration, deadband handling, and pedal release shaping.";
    evidence(end + 1) = sprintf('Cruise torque leakage share is %.1f%% above %.1f Nm.', ...
        100 * RCA_FractionTrue(cruiseLeakageMask, cruisePhase), leakageTorqueThreshold);
end
if any(validGear) && 100 * RCA_FractionTrue(driveNearLimitMask & shiftMask, shiftMask & driveMask) > 15
    recs(end + 1) = "Review shift-aware feedforward torque shaping so the controller does not repeatedly demand near-limit positive torque while the actual gear is changing.";
    evidence(end + 1) = sprintf('Near positive limit during shift share is %.1f%%.', ...
        100 * RCA_FractionTrue(driveNearLimitMask & shiftMask, shiftMask & driveMask));
end
if any(validGear) && 100 * RCA_FractionTrue(regenNearLimitMask & shiftMask, shiftMask & regenMask) > 15
    recs(end + 1) = "Review shift-aware recuperation shaping and gear-state handover so regen demand does not crowd the negative torque limit during gear transitions.";
    evidence(end + 1) = sprintf('Near regen limit during shift share is %.1f%%.', ...
        100 * RCA_FractionTrue(regenNearLimitMask & shiftMask, shiftMask & regenMask));
end
if any(validGear) && gearShiftCount > 0 && mean(abs(totalDemand(shiftMask)), 'omitnan') > mean(abs(totalDemand(~shiftMask)), 'omitnan') * 1.25
    recs(end + 1) = "Check whether actual-gear or gear-ratio feedforward terms are amplifying torque demand around shifts; torque command magnitude rises materially during shift activity.";
    evidence(end + 1) = sprintf('Mean absolute torque during shift activity is %.1f Nm versus %.1f Nm away from shifts.', ...
        mean(abs(totalDemand(shiftMask)), 'omitnan'), mean(abs(totalDemand(~shiftMask)), 'omitnan'));
end

if isfinite(accelNonlinearity) && accelNonlinearity > config.Thresholds.ControllerMappingNonlinearityWarn
    recs(end + 1) = "Review accelerator-to-torque map shape, clipping, and filtering; the mapping nonlinearity index indicates the commanded torque is not well represented by a simple monotonic pedal gain.";
    evidence(end + 1) = sprintf('Accelerator mapping nonlinearity index is %.2f.', accelNonlinearity);
end
if isfinite(brakeNonlinearity) && brakeNonlinearity > config.Thresholds.ControllerMappingNonlinearityWarn
    recs(end + 1) = "Review brake-to-regen map shape and blending logic; high nonlinearity often indicates regen clipping, friction substitution, or discontinuous brake blending.";
    evidence(end + 1) = sprintf('Brake regen mapping nonlinearity index is %.2f.', brakeNonlinearity);
end
if any(validDischargePowerLimit) && RCA_FractionTrue(nearDischargePowerLimit, validDischargePowerLimit) > 0.05
    recs(end + 1) = "Acceleration shortfall should be reviewed with BMS discharge power limits; the motor electrical power frequently approaches the battery discharge envelope.";
    evidence(end + 1) = sprintf('Near BMS discharge power limit share is %.1f%%.', 100 * RCA_FractionTrue(nearDischargePowerLimit, validDischargePowerLimit));
end
if any(validChargePowerLimit) && RCA_FractionTrue(nearChargePowerLimit, validChargePowerLimit) > 0.05
    recs(end + 1) = "Regeneration shortfall should be reviewed with BMS charge power limits and SoC; the motor recuperation power frequently approaches battery charge acceptance.";
    evidence(end + 1) = sprintf('Near BMS charge power limit share is %.1f%%.', 100 * RCA_FractionTrue(nearChargePowerLimit, validChargePowerLimit));
end
if any(signMismatchMask)
    recs(end + 1) = "Investigate demand-to-actual torque sign mismatch during mode handover; this can indicate lag, sign convention mismatch, or delayed torque reversal.";
    evidence(end + 1) = sprintf('Torque sign mismatch share is %.2f%% of valid actual-torque samples.', 100 * RCA_FractionTrue(signMismatchMask, validActual));
end

constraintEvents = localBuildConstraintEventTable(t, totalDemand, totalActual, torqueError, accPedal, brkPedal, gearNum, battSoc, ...
    driveNearLimitMask, regenNearLimitMask, nearDischargePowerLimit, nearChargePowerLimit, shiftMask, config);
localSafeWriteTable(constraintEvents, fullfile(outputPaths.Tables, 'PowerTrainController_ConstraintEvents.csv'));
if ~isempty(constraintEvents)
    topEventCount = min(3, height(constraintEvents));
    for iEvent = 1:topEventCount
        summary(end + 1) = sprintf('Powertrain controller event %d: %s from %.1f s to %.1f s. Likely cause: %s. Confidence: %s.', ...
            constraintEvents.EventID(iEvent), constraintEvents.EventType(iEvent), constraintEvents.StartTime_s(iEvent), ...
            constraintEvents.EndTime_s(iEvent), constraintEvents.LikelyCause(iEvent), constraintEvents.ConfidenceNote(iEvent));
    end
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'PowerTrainController');
plotFiles = localAppendPlotFile(plotFiles, localPlotCommandOverview(figureFolder, t, accPedal, brkPedal, totalDemand, totalActual, posLimit, negLimit, driveNearLimitMask, regenNearLimitMask, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotMotorSplit(figureFolder, t, em1Demand, em2Demand, totalDemand, accelPhase, brakePhase, cruisePhase, splitImbalanceDrive, splitImbalanceRegen, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotPedalResponse(figureFolder, accPedal, brkPedal, totalDemand, driveNearLimitMask, regenNearLimitMask, accelPhase, brakePhase, cruisePhase, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotGearContext(figureFolder, t, gearNum, gearRatio, totalDemand, posLimit, negLimit, shiftMask, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotSystemOverview(figureFolder, t, accPedal, brkPedal, totalDemand, totalActual, posLimit, negLimit, motorElecPwr, battDischargeLimit, battChargeLimit, battSoc, gearNum, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotTorqueTrackingQuality(figureFolder, totalDemand, totalActual, torqueError, gearNum, battSoc, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotPowerUtilization(figureFolder, t, vehSpeed, battSoc, gearNum, drivePowerUse, regenPowerUse, nearDischargePowerLimit, nearChargePowerLimit, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotConstraintEvents(figureFolder, t, totalDemand, totalActual, torqueError, constraintEvents, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.FigureFiles = plotFiles;
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Power Train Controller", recs, evidence);
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

function rate = localSignalRate(signal, dt)
signal = double(signal(:));
dt = double(dt(:));
rate = NaN(size(signal));
if isempty(signal)
    return;
end
if numel(signal) == 1
    rate(1) = 0;
    return;
end
rate(1) = 0;
deltaTime = dt(2:end);
deltaTime(~isfinite(deltaTime) | deltaTime <= 0) = NaN;
rate(2:end) = diff(signal) ./ deltaTime;
end

function value = localRmsFinite(signal, mask)
signal = double(signal(:));
if nargin < 2 || isempty(mask)
    mask = true(size(signal));
else
    mask = logical(mask(:));
    if numel(mask) ~= numel(signal)
        mask = true(size(signal));
    end
end
valid = mask & isfinite(signal);
if ~any(valid)
    value = NaN;
    return;
end
value = sqrt(mean(signal(valid).^2, 'omitnan'));
end

function [gain, corrValue, nonlinearityIndex, deadbandEstimate] = localPedalTorqueMappingStats(pedal, torqueMagnitude, mask, torqueDeadband)
pedal = double(pedal(:));
torqueMagnitude = double(torqueMagnitude(:));
mask = logical(mask(:));
valid = mask & isfinite(pedal) & isfinite(torqueMagnitude);
gain = NaN;
corrValue = NaN;
nonlinearityIndex = NaN;
deadbandEstimate = NaN;
if sum(valid) < 3
    return;
end
x = pedal(valid);
y = torqueMagnitude(valid);
active = y > torqueDeadband;
if any(active)
    deadbandEstimate = min(x(active), [], 'omitnan');
end
xCentered = x - mean(x, 'omitnan');
yCentered = y - mean(y, 'omitnan');
denGain = sum(xCentered.^2, 'omitnan');
denCorr = sqrt(sum(xCentered.^2, 'omitnan') * sum(yCentered.^2, 'omitnan'));
if denGain > eps
    gain = sum(xCentered .* yCentered, 'omitnan') / denGain;
    intercept = mean(y, 'omitnan') - gain * mean(x, 'omitnan');
    residual = y - (gain .* x + intercept);
    nonlinearityIndex = sqrt(mean(residual.^2, 'omitnan')) / max(RCA_Percentile(abs(y), 95), eps);
end
if denCorr > eps
    corrValue = sum(xCentered .* yCentered, 'omitnan') / denCorr;
end
end

function count = localCommandReversalCount(totalDemand, deadband)
state = zeros(size(totalDemand));
state(totalDemand > deadband) = 1;
state(totalDemand < -deadband) = -1;
state = state(state ~= 0);
if numel(state) <= 1
    count = 0;
    return;
end
count = sum(state(2:end) ~= state(1:end - 1));
end

function localSafeWriteTable(tableValue, filePath)
try
    writetable(tableValue, filePath);
catch
end
end

function imbalancePct = localSplitImbalance(em1, em2, totalDemand, activeMask)
imbalancePct = NaN(size(totalDemand));
valid = activeMask & isfinite(em1) & isfinite(em2) & abs(totalDemand) > eps;
imbalancePct(valid) = 100 * abs(em1(valid) - em2(valid)) ./ max(abs(totalDemand(valid)), eps);
end

function sharePct = localMotorNearLimit(demand, limitSignal, activeMask, usageFraction, positiveMode)
sharePct = NaN;
if positiveMode
    valid = activeMask & isfinite(demand) & isfinite(limitSignal) & limitSignal > 0;
    sharePct = 100 * RCA_FractionTrue(demand >= usageFraction .* limitSignal, valid);
else
    valid = activeMask & isfinite(demand) & isfinite(limitSignal) & abs(limitSignal) > 0;
    sharePct = 100 * RCA_FractionTrue(abs(demand) >= usageFraction .* abs(limitSignal), valid);
end
end

function shiftEvents = localShiftEvents(gearNum)
shiftEvents = false(size(gearNum));
if numel(gearNum) < 2
    return;
end
validStep = isfinite(gearNum(2:end)) & isfinite(gearNum(1:end-1));
shiftEvents(2:end) = validStep & abs(diff(gearNum)) > 0.05;
end

function shiftMask = localShiftMask(gearNum)
shiftMask = localShiftEvents(gearNum);
if any(shiftMask)
    shiftMask = shiftMask | [shiftMask(2:end); false] | [false; shiftMask(1:end-1)];
end
end

function [gearList, driveMean, regenMean, driveShare] = localGearDemandStats(gearNum, totalDemand, driveMask, regenMask)
validGear = isfinite(gearNum);
gearList = unique(round(gearNum(validGear)));
gearList = gearList(:)';
driveMean = NaN(size(gearList));
regenMean = NaN(size(gearList));
driveShare = NaN(size(gearList));
for iGear = 1:numel(gearList)
    gearMask = validGear & round(gearNum) == gearList(iGear);
    driveMean(iGear) = mean(totalDemand(gearMask & driveMask), 'omitnan');
    regenMean(iGear) = mean(max(-totalDemand(gearMask & regenMask), 0), 'omitnan');
    driveShare(iGear) = 100 * RCA_FractionTrue(gearMask & driveMask, driveMask);
end
end

function [dominantGear, dominantShare] = localDominantGear(gearNum, activeMask)
dominantGear = NaN;
dominantShare = NaN;
valid = isfinite(gearNum) & activeMask;
if ~any(valid)
    return;
end
gearVals = round(gearNum(valid));
gearList = unique(gearVals);
gearCounts = zeros(size(gearList));
for iGear = 1:numel(gearList)
    gearCounts(iGear) = sum(gearVals == gearList(iGear));
end
[peakCount, idx] = max(gearCounts);
dominantGear = gearList(idx);
dominantShare = 100 * peakCount / sum(gearCounts);
end

function eventTable = localBuildConstraintEventTable(t, totalDemand, totalActual, torqueError, accPedal, brkPedal, gearNum, battSoc, ...
    driveNearLimitMask, regenNearLimitMask, nearDischargePowerLimit, nearChargePowerLimit, shiftMask, config)
eventTable = localEmptyConstraintEventTable();
if isempty(t)
    return;
end
trackingPoor = isfinite(torqueError) & abs(torqueError) >= config.Thresholds.ControllerTorqueTrackingWarn_Nm;
eventMask = driveNearLimitMask | regenNearLimitMask | nearDischargePowerLimit | nearChargePowerLimit | ...
    (shiftMask & trackingPoor) | trackingPoor;
if ~any(eventMask)
    return;
end
startIdx = find(eventMask & ~[false; eventMask(1:end-1)]);
endIdx = find(eventMask & ~[eventMask(2:end); false]);
rows = cell(0, 16);
for iEvent = 1:numel(startIdx)
    idx = startIdx(iEvent):endIdx(iEvent);
    duration_s = max(t(endIdx(iEvent)) - t(startIdx(iEvent)), 0);
    if duration_s < config.Thresholds.MinEventDuration_s
        continue;
    end
    meanDemand = mean(totalDemand(idx), 'omitnan');
    meanActual = mean(totalActual(idx), 'omitnan');
    meanAbsError = mean(abs(torqueError(idx)), 'omitnan');
    eventType = localConstraintEventType(idx, driveNearLimitMask, regenNearLimitMask, nearDischargePowerLimit, nearChargePowerLimit, shiftMask, torqueError, config);
    [likelyCause, confidenceNote, recommendation] = localInterpretConstraintEvent(eventType, meanDemand, meanAbsError, ...
        mean(accPedal(idx), 'omitnan'), mean(brkPedal(idx), 'omitnan'), mean(gearNum(idx), 'omitnan'), mean(battSoc(idx), 'omitnan'));
    severity = meanAbsError / max(config.Thresholds.ControllerTorqueTrackingWarn_Nm, eps) + ...
        0.35 * RCA_FractionTrue(driveNearLimitMask(idx) | regenNearLimitMask(idx), true(numel(idx), 1)) + ...
        0.35 * RCA_FractionTrue(nearDischargePowerLimit(idx) | nearChargePowerLimit(idx), true(numel(idx), 1)) + ...
        0.20 * RCA_FractionTrue(shiftMask(idx), true(numel(idx), 1));
    rows(end + 1, :) = {size(rows, 1) + 1, t(startIdx(iEvent)), t(endIdx(iEvent)), duration_s, ...
        string(eventType), meanDemand, meanActual, meanAbsError, max(abs(torqueError(idx)), [], 'omitnan'), ...
        mean(accPedal(idx), 'omitnan'), mean(brkPedal(idx), 'omitnan'), mean(gearNum(idx), 'omitnan'), ...
        severity, string(likelyCause), string(confidenceNote), string(recommendation)}; %#ok<AGROW>
end
if isempty(rows)
    return;
end
eventTable = cell2table(rows, 'VariableNames', {'EventID', 'StartTime_s', 'EndTime_s', 'Duration_s', ...
    'EventType', 'MeanDemandTorque_Nm', 'MeanActualTorque_Nm', 'MeanAbsTorqueError_Nm', ...
    'PeakAbsTorqueError_Nm', 'MeanAccelPedal_pct', 'MeanBrakePedal_pct', 'MeanGear', ...
    'Severity', 'LikelyCause', 'ConfidenceNote', 'RecommendedAction'});
eventTable = sortrows(eventTable, {'Severity', 'PeakAbsTorqueError_Nm'}, {'descend', 'descend'});
eventTable.EventID = (1:height(eventTable))';
end

function eventTable = localEmptyConstraintEventTable()
eventTable = cell2table(cell(0, 16), 'VariableNames', {'EventID', 'StartTime_s', 'EndTime_s', 'Duration_s', ...
    'EventType', 'MeanDemandTorque_Nm', 'MeanActualTorque_Nm', 'MeanAbsTorqueError_Nm', ...
    'PeakAbsTorqueError_Nm', 'MeanAccelPedal_pct', 'MeanBrakePedal_pct', 'MeanGear', ...
    'Severity', 'LikelyCause', 'ConfidenceNote', 'RecommendedAction'});
end

function eventType = localConstraintEventType(idx, driveNearLimitMask, regenNearLimitMask, nearDischargePowerLimit, nearChargePowerLimit, shiftMask, torqueError, config)
if RCA_FractionTrue(nearDischargePowerLimit(idx), true(numel(idx), 1)) > 0.4
    eventType = "Battery discharge power-limited drive";
elseif RCA_FractionTrue(nearChargePowerLimit(idx), true(numel(idx), 1)) > 0.4
    eventType = "Battery charge power-limited regen";
elseif RCA_FractionTrue(driveNearLimitMask(idx), true(numel(idx), 1)) > 0.4
    eventType = "Positive torque envelope saturation";
elseif RCA_FractionTrue(regenNearLimitMask(idx), true(numel(idx), 1)) > 0.4
    eventType = "Negative torque envelope saturation";
elseif RCA_FractionTrue(shiftMask(idx), true(numel(idx), 1)) > 0.2
    eventType = "Shift-related torque disturbance";
elseif mean(abs(torqueError(idx)), 'omitnan') >= config.Thresholds.ControllerTorqueTrackingWarn_Nm
    eventType = "Demand-to-actual torque tracking error";
else
    eventType = "Mixed controller constraint event";
end
end

function [likelyCause, confidenceNote, recommendation] = localInterpretConstraintEvent(eventType, meanDemand, meanAbsError, meanAccel, meanBrake, meanGear, meanSoc)
likelyCause = "Mixed controller, actuator, or limit interaction";
confidenceNote = "Low-medium confidence because multiple constraint indicators overlap.";
recommendation = "Review aligned pedal, torque demand, actual torque, limits, gear, and battery state for the event.";
switch string(eventType)
    case "Battery discharge power-limited drive"
        likelyCause = "BMS discharge power limit constraining propulsion request";
        confidenceNote = "Medium-high confidence because drive power use is near the battery discharge envelope.";
        recommendation = "Do not treat acceleration shortfall as pedal-map tuning only; review discharge power-limit calibration and battery capability.";
    case "Battery charge power-limited regen"
        likelyCause = "BMS charge power limit constraining recuperation request";
        confidenceNote = "Medium-high confidence because regen power use is near the battery charge acceptance envelope.";
        recommendation = "Review charge power limit, SoC dependency, and regen-to-friction blending.";
    case "Positive torque envelope saturation"
        likelyCause = "Motor/controller positive torque envelope reached";
        confidenceNote = "Medium confidence because commanded torque is near the available positive torque limit.";
        recommendation = "Review motor torque-speed envelope, gear state, and requested torque ramp before increasing driver demand gain.";
    case "Negative torque envelope saturation"
        likelyCause = "Motor/controller recuperation torque envelope reached";
        confidenceNote = "Medium confidence because commanded negative torque is near the available recuperation limit.";
        recommendation = "Review regen limit scheduling and brake blending, especially at low speed or high SoC.";
    case "Shift-related torque disturbance"
        likelyCause = "Gear-state handover or gear-ratio feedforward disturbance";
        confidenceNote = "Medium confidence because the event overlaps gear transition samples.";
        recommendation = "Review shift-aware torque shaping and torque handover between ratio states.";
    case "Demand-to-actual torque tracking error"
        likelyCause = "Actuator response lag, clipping, or downstream torque delivery limitation";
        confidenceNote = "Medium confidence because torque error is high without a stronger classified limit signature.";
        recommendation = "Review demand-to-actual torque delay, rate limits, and eDrive tracking around the event.";
end
if isfinite(meanSoc)
    confidenceNote = confidenceNote + sprintf(' Mean SoC during event is %.1f%%.', meanSoc);
end
if isfinite(meanGear)
    recommendation = recommendation + sprintf(' Mean gear during event is %.1f.', meanGear);
end
if isfinite(meanDemand) && isfinite(meanAbsError)
    confidenceNote = confidenceNote + sprintf(' Mean demand %.1f Nm, mean |error| %.1f Nm.', meanDemand, meanAbsError);
end
if isfinite(meanAccel) && meanAccel > 50
    recommendation = recommendation + " High accelerator demand indicates this event affects propulsion feel.";
elseif isfinite(meanBrake) && meanBrake > 20
    recommendation = recommendation + " Brake demand indicates this event affects recuperation/brake blending.";
end
end

function plotFile = localPlotCommandOverview(outputFolder, t, accPedal, brkPedal, totalDemand, totalActual, posLimit, negLimit, driveNearLimitMask, regenNearLimitMask, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, accPedal, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, brkPedal, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Powertrain Controller Inputs');
ylabel('Pedal (%)');
legend({'Accelerator pedal', 'Brake pedal'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
hLegend = gobjects(0, 1);
legendText = strings(0, 1);
hLegend(end + 1) = plot(t, totalDemand, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
legendText(end + 1) = "Demand torque";
hLegend(end + 1) = plot(t, totalActual, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
legendText(end + 1) = "Actual torque";
if any(isfinite(posLimit))
    hLegend(end + 1) = plot(t, posLimit, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
    legendText(end + 1) = "Max available torque";
end
if any(isfinite(negLimit))
    hLegend(end + 1) = plot(t, negLimit, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
    legendText(end + 1) = "Min available torque";
end
hLegend(end + 1) = plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
legendText(end + 1) = "Zero line";
title('Total Torque Demand, Delivery, and Available Envelope');
ylabel('Torque (Nm)');
legend(hLegend, cellstr(legendText), 'Location', 'best');
grid on;

subplot(3, 1, 3);
driveUtil = NaN(size(totalDemand));
regenUtil = NaN(size(totalDemand));
validDrive = isfinite(posLimit) & posLimit > 0 & totalDemand > 0;
validRegen = isfinite(negLimit) & abs(negLimit) > 0 & totalDemand < 0;
driveUtil(validDrive) = 100 * totalDemand(validDrive) ./ posLimit(validDrive);
regenUtil(validRegen) = 100 * abs(totalDemand(validRegen)) ./ abs(negLimit(validRegen));
hLegend = gobjects(0, 1);
legendText = strings(0, 1);
hLegend(end + 1) = plot(t, driveUtil, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
legendText(end + 1) = "Drive limit usage";
hLegend(end + 1) = plot(t, regenUtil, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
legendText(end + 1) = "Regen limit usage";
if any(driveNearLimitMask)
    hLegend(end + 1) = plot(t(driveNearLimitMask), driveUtil(driveNearLimitMask), 'o', 'Color', config.Plot.Colors.Warning, 'MarkerSize', 4);
    legendText(end + 1) = "Near positive limit";
end
if any(regenNearLimitMask)
    hLegend(end + 1) = plot(t(regenNearLimitMask), regenUtil(regenNearLimitMask), 'o', 'Color', config.Plot.Colors.Auxiliary, 'MarkerSize', 4);
    legendText(end + 1) = "Near regen limit";
end
title('Available Torque Utilization');
xlabel('Time (s)');
ylabel('Limit usage (%)');
legend(hLegend, cellstr(legendText), 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_CommandOverview', config));
close(fig);
end

function plotFile = localPlotMotorSplit(outputFolder, t, em1Demand, em2Demand, totalDemand, accelPhase, brakePhase, cruisePhase, splitImbalanceDrive, splitImbalanceRegen, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, em1Demand, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, em2Demand, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Motor 1 and Motor 2 Demand Torque');
ylabel('Torque (Nm)');
legend({'Motor 1 demand', 'Motor 2 demand', 'Zero line'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, splitImbalanceDrive, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, splitImbalanceRegen, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.ControllerSplitImbalance_pct, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Motor Command Split Imbalance');
ylabel('Imbalance (%)');
legend({'Drive imbalance', 'Regen imbalance', 'Review threshold'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
phaseCode = NaN(size(totalDemand));
phaseCode(accelPhase) = 1;
phaseCode(cruisePhase) = 2;
phaseCode(brakePhase) = 3;
valid = isfinite(em1Demand) & isfinite(em2Demand) & isfinite(phaseCode);
scatter(em1Demand(valid), em2Demand(valid), 12, phaseCode(valid), 'filled');
colormap([config.Plot.Colors.Demand; config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
cb = colorbar;
cb.Ticks = [1 2 3];
cb.TickLabels = {'Acceleration', 'Cruise', 'Braking'};
title('Motor 1 vs Motor 2 Demand Torque');
xlabel('Motor 1 demand torque (Nm)');
ylabel('Motor 2 demand torque (Nm)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_MotorSplit', config));
close(fig);
end

function plotFile = localPlotPedalResponse(outputFolder, accPedal, brkPedal, totalDemand, driveNearLimitMask, regenNearLimitMask, accelPhase, brakePhase, cruisePhase, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

driveDemand = max(totalDemand, 0);
regenDemand = max(-totalDemand, 0);
phaseMean = [mean(driveDemand(accelPhase), 'omitnan'), mean(regenDemand(brakePhase), 'omitnan'), mean(abs(totalDemand(cruisePhase)), 'omitnan')];
phaseShare = [100 * RCA_FractionTrue(totalDemand > config.Thresholds.ControllerTorqueDeadband_Nm, accelPhase), ...
    100 * RCA_FractionTrue(totalDemand < -config.Thresholds.ControllerTorqueDeadband_Nm, brakePhase), ...
    100 * RCA_FractionTrue(abs(totalDemand) > config.Thresholds.ControllerLeakageTorque_Nm, cruisePhase)];

subplot(2, 2, 1);
driveUtilColor = double(driveNearLimitMask);
scatter(accPedal(isfinite(accPedal) & driveDemand > 0), driveDemand(isfinite(accPedal) & driveDemand > 0), 12, driveUtilColor(isfinite(accPedal) & driveDemand > 0), 'filled');
colormap(gca, [config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
title('Accelerator Pedal vs Drive Torque Request');
xlabel('Accelerator pedal (%)');
ylabel('Drive torque demand (Nm)');
grid on;

subplot(2, 2, 2);
regenUtilColor = double(regenNearLimitMask);
scatter(brkPedal(isfinite(brkPedal) & regenDemand > 0), regenDemand(isfinite(brkPedal) & regenDemand > 0), 12, regenUtilColor(isfinite(brkPedal) & regenDemand > 0), 'filled');
colormap(gca, [config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
title('Brake Pedal vs Regen Torque Request');
xlabel('Brake pedal (%)');
ylabel('Regen torque demand magnitude (Nm)');
grid on;

subplot(2, 2, 3);
bar(categorical({'Acceleration', 'Braking', 'Cruise'}), phaseMean, 'FaceColor', config.Plot.Colors.Vehicle);
title('Mean Torque Demand by Phase');
ylabel('Torque / torque magnitude (Nm)');
grid on;

subplot(2, 2, 4);
bar(categorical({'Accel drive share', 'Brake regen share', 'Cruise leakage share'}), phaseShare, 'FaceColor', config.Plot.Colors.Motor);
title('Phase Command Quality Shares');
ylabel('Share (%)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_PedalResponse', config));
close(fig);
end

function plotFile = localPlotGearContext(outputFolder, t, gearNum, gearRatio, totalDemand, posLimit, negLimit, shiftMask, config)
plotFile = "";
if ~any(isfinite(gearNum)) && ~any(isfinite(gearRatio))
    return;
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(4, 1, 1);
if any(isfinite(gearNum))
    hLegend = gobjects(0, 1);
    legendText = strings(0, 1);
    hLegend(end + 1) = stairs(t, gearNum, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    legendText(end + 1) = "Actual gear";
    if any(shiftMask)
        hLegend(end + 1) = plot(t(shiftMask), gearNum(shiftMask), 'o', 'Color', config.Plot.Colors.Warning, 'MarkerSize', 4);
        legendText(end + 1) = "Shift activity";
    end
    ylabel('Gear (-)');
    title('Actual Gear Number and Shift Activity');
    legend(hLegend, cellstr(legendText), 'Location', 'best');
else
    plot(t, NaN(size(t)));
    title('Actual Gear Number and Shift Activity');
    ylabel('Gear (-)');
end
grid on;

subplot(4, 1, 2);
if any(isfinite(gearRatio))
    plot(t, gearRatio, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
end
title('Actual Gear Ratio');
ylabel('Ratio (-)');
grid on;

subplot(4, 1, 3);
hLegend = gobjects(0, 1);
legendText = strings(0, 1);
hLegend(end + 1) = plot(t, totalDemand, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
legendText(end + 1) = "Demand torque";
if any(isfinite(posLimit))
    hLegend(end + 1) = plot(t, posLimit, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
    legendText(end + 1) = "Max available torque";
end
if any(isfinite(negLimit))
    hLegend(end + 1) = plot(t, negLimit, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
    legendText(end + 1) = "Min available torque";
end
if any(shiftMask)
    hLegend(end + 1) = plot(t(shiftMask), totalDemand(shiftMask), 'o', 'Color', config.Plot.Colors.Vehicle, 'MarkerSize', 4);
    legendText(end + 1) = "Shift activity";
end
hLegend(end + 1) = plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
legendText(end + 1) = "Zero line";
title('Torque Demand with Shift Context');
ylabel('Torque (Nm)');
legend(hLegend, cellstr(legendText), 'Location', 'best');
grid on;

subplot(4, 1, 4);
validScatter = isfinite(gearRatio) & isfinite(totalDemand);
scatter(gearRatio(validScatter), totalDemand(validScatter), 12, double(shiftMask(validScatter)), 'filled');
colormap(gca, [config.Plot.Colors.Vehicle; config.Plot.Colors.Warning]);
cb = colorbar;
cb.Ticks = [0 1];
cb.TickLabels = {'Steady gear', 'Shift activity'};
title('Torque Demand vs Actual Gear Ratio');
xlabel('Actual gear ratio (-)');
ylabel('Torque demand (Nm)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_GearContext', config));
close(fig);
end

function plotFile = localPlotSystemOverview(outputFolder, t, accPedal, brkPedal, totalDemand, totalActual, posLimit, negLimit, motorElecPwr, battDischargeLimit, battChargeLimit, battSoc, gearNum, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', [100 100 1400 900]);

subplot(5, 1, 1);
plot(t, accPedal, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, brkPedal, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Driver Intent Inputs');
ylabel('Pedal [%]');
legend({'Accelerator', 'Brake'}, 'Location', 'best');
grid on;

subplot(5, 1, 2);
plot(t, totalDemand, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, totalActual, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
if any(isfinite(posLimit))
    plot(t, posLimit, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
end
if any(isfinite(negLimit))
    plot(t, negLimit, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
end
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Torque Demand, Actual Torque, and Controller Torque Envelope');
ylabel('Torque [Nm]');
legend({'Demand', 'Actual', 'Positive limit', 'Negative limit', 'Zero'}, 'Location', 'best');
grid on;

subplot(5, 1, 3);
plot(t, max(motorElecPwr, 0), 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, -max(-motorElecPwr, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
if any(isfinite(battDischargeLimit))
    plot(t, battDischargeLimit, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
end
if any(isfinite(battChargeLimit))
    plot(t, -battChargeLimit, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
end
title('Motor Electrical Power Against BMS Power Limits');
ylabel('Power [kW]');
legend({'Drive power', 'Regen power', 'Discharge limit', 'Charge limit'}, 'Location', 'best');
grid on;

subplot(5, 1, 4);
yyaxis left;
plot(t, battSoc, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
ylabel('SoC [%]');
yyaxis right;
stairs(t, gearNum, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
ylabel('Gear [-]');
title('Battery State and Actual Gear Context');
grid on;

subplot(5, 1, 5);
torqueError = totalDemand - totalActual;
plot(t, torqueError, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', [0.55 0.55 0.55]);
yline(-config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', [0.55 0.55 0.55]);
title('Demand-to-Actual Torque Error');
ylabel('Error [Nm]');
xlabel('Time [s]');
grid on;

sgtitle('Powertrain Controller System Overview');
plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_SystemOverview', config));
close(fig);
end

function plotFile = localPlotTorqueTrackingQuality(outputFolder, totalDemand, totalActual, torqueError, gearNum, battSoc, config)
plotFile = "";
if ~any(isfinite(totalDemand)) || ~any(isfinite(totalActual))
    return;
end
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid = isfinite(totalDemand) & isfinite(totalActual);
scatter(totalDemand(valid), totalActual(valid), 12, config.Plot.Colors.Vehicle, 'filled');
hold on;
lims = [min([totalDemand(valid); totalActual(valid)], [], 'omitnan'), max([totalDemand(valid); totalActual(valid)], [], 'omitnan')];
if all(isfinite(lims)) && lims(2) > lims(1)
    plot(lims, lims, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.1);
end
xlabel('Demand torque [Nm]');
ylabel('Actual torque [Nm]');
title('Demand vs Actual Torque');
grid on;

subplot(2, 2, 2);
validError = torqueError(isfinite(torqueError));
if ~isempty(validError)
    histogram(validError, 40, 'FaceColor', config.Plot.Colors.Vehicle, 'EdgeColor', 'none');
end
xline(config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', [0.55 0.55 0.55]);
xline(-config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', [0.55 0.55 0.55]);
xlabel('Torque error [Nm]');
ylabel('Samples');
title('Torque Error Distribution');
grid on;

subplot(2, 2, 3);
validGear = isfinite(gearNum) & isfinite(torqueError);
if any(validGear)
    gearValues = unique(round(gearNum(validGear)));
    gearMae = NaN(size(gearValues));
    for iGear = 1:numel(gearValues)
        gearMae(iGear) = mean(abs(torqueError(validGear & round(gearNum) == gearValues(iGear))), 'omitnan');
    end
    bar(gearValues, gearMae, 'FaceColor', config.Plot.Colors.Gear);
end
xlabel('Gear [-]');
ylabel('Mean |error| [Nm]');
title('Torque Tracking by Gear');
grid on;

subplot(2, 2, 4);
validSoc = isfinite(battSoc) & isfinite(torqueError);
if any(validSoc)
    scatter(battSoc(validSoc), abs(torqueError(validSoc)), 12, config.Plot.Colors.Battery, 'filled');
end
xlabel('Battery SoC [%]');
ylabel('|Torque error| [Nm]');
title('Torque Error vs Battery State');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_TorqueTrackingQuality', config));
close(fig);
end

function plotFile = localPlotPowerUtilization(outputFolder, t, vehSpeed, battSoc, gearNum, drivePowerUse, regenPowerUse, nearDischargePowerLimit, nearChargePowerLimit, config)
plotFile = "";
if ~any(isfinite(drivePowerUse)) && ~any(isfinite(regenPowerUse))
    return;
end
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
plot(t, drivePowerUse * 100, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, regenPowerUse * 100, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.LimitUsageFraction * 100, '--', 'Color', [0.55 0.55 0.55]);
title('BMS Power Limit Utilization');
xlabel('Time [s]');
ylabel('Utilization [%]');
legend({'Drive/discharge', 'Regen/charge', 'Near-limit threshold'}, 'Location', 'best');
grid on;

subplot(2, 2, 2);
validSpeedDrive = isfinite(vehSpeed) & isfinite(drivePowerUse);
validSpeedRegen = isfinite(vehSpeed) & isfinite(regenPowerUse);
hold on;
scatter(vehSpeed(validSpeedDrive), drivePowerUse(validSpeedDrive) * 100, 12, config.Plot.Colors.Demand, 'filled');
scatter(vehSpeed(validSpeedRegen), regenPowerUse(validSpeedRegen) * 100, 12, config.Plot.Colors.Battery, 'filled');
xlabel('Vehicle speed [km/h]');
ylabel('Utilization [%]');
title('Power Utilization vs Speed');
legend({'Drive', 'Regen'}, 'Location', 'best');
grid on;

subplot(2, 2, 3);
validSocDrive = isfinite(battSoc) & isfinite(drivePowerUse);
validSocRegen = isfinite(battSoc) & isfinite(regenPowerUse);
hold on;
scatter(battSoc(validSocDrive), drivePowerUse(validSocDrive) * 100, 12, double(nearDischargePowerLimit(validSocDrive)), 'filled');
scatter(battSoc(validSocRegen), regenPowerUse(validSocRegen) * 100, 12, double(nearChargePowerLimit(validSocRegen)) + 2, 'filled');
xlabel('Battery SoC [%]');
ylabel('Utilization [%]');
title('Power Utilization vs SoC');
grid on;

subplot(2, 2, 4);
validGear = isfinite(gearNum) & (isfinite(drivePowerUse) | isfinite(regenPowerUse));
if any(validGear)
    gearValues = unique(round(gearNum(validGear)));
    driveMean = NaN(size(gearValues));
    regenMean = NaN(size(gearValues));
    for iGear = 1:numel(gearValues)
        gearMask = round(gearNum) == gearValues(iGear);
        driveMean(iGear) = mean(drivePowerUse(gearMask), 'omitnan') * 100;
        regenMean(iGear) = mean(regenPowerUse(gearMask), 'omitnan') * 100;
    end
    bar(gearValues, [driveMean(:), regenMean(:)]);
    legend({'Drive', 'Regen'}, 'Location', 'best');
end
xlabel('Gear [-]');
ylabel('Mean utilization [%]');
title('Power Utilization by Gear');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_PowerUtilization', config));
close(fig);
end

function plotFile = localPlotConstraintEvents(outputFolder, t, totalDemand, totalActual, torqueError, eventTable, config)
plotFile = "";
if isempty(eventTable) || height(eventTable) == 0
    return;
end
eventTable = eventTable(1:min(5, height(eventTable)), :);
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 1, 1);
plot(t, totalDemand, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, totalActual, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
localShadeIntervals(gca, eventTable.StartTime_s, eventTable.EndTime_s, [0.98 0.88 0.78], 0.65);
title('Top Constraint Events on Torque Timeline');
ylabel('Torque [Nm]');
legend({'Demand', 'Actual', 'Top events'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
plot(t, torqueError, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', [0.55 0.55 0.55]);
yline(-config.Thresholds.ControllerTorqueTrackingWarn_Nm, '--', 'Color', [0.55 0.55 0.55]);
localShadeIntervals(gca, eventTable.StartTime_s, eventTable.EndTime_s, [0.98 0.88 0.78], 0.65);
title('Torque Error During Top Constraint Events');
ylabel('Torque error [Nm]');
xlabel('Time [s]');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PowerTrainController_ConstraintEvents', config));
close(fig);
end

function patchHandle = localShadeIntervals(axisHandle, startTimes, endTimes, colorValue, alphaValue)
patchHandle = gobjects(1, 1);
startTimes = double(startTimes(:));
endTimes = double(endTimes(:));
valid = isfinite(startTimes) & isfinite(endTimes) & endTimes >= startTimes;
startTimes = startTimes(valid);
endTimes = endTimes(valid);
if isempty(startTimes)
    return;
end
yLimits = ylim(axisHandle);
for iInt = 1:numel(startTimes)
    currentPatch = patch(axisHandle, [startTimes(iInt) endTimes(iInt) endTimes(iInt) startTimes(iInt)], ...
        [yLimits(1) yLimits(1) yLimits(2) yLimits(2)], colorValue, ...
        'FaceAlpha', alphaValue, 'EdgeColor', 'none');
    if iInt == 1
        patchHandle = currentPatch;
    end
end
uistack(patchHandle, 'bottom');
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
