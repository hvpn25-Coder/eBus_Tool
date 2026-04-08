function result = Analyze_PowerTrainController(analysisData, outputPaths, config)
% Analyze_PowerTrainController  Pedal-to-torque command quality and saturation analysis.

result = localInitResult("POWER TRAIN CONTROLLER", ...
    {'acc_pdl', 'brk_pdl', 'emot1_dem_trq', 'emot2_dem_trq'}, ...
    {'emot1_act_trq', 'emot2_act_trq', 'max_emot1_dem_trq', 'max_emot2_dem_trq', ...
    'min_emot1_dem_trq', 'min_emot2_dem_trq', 'gr_num', 'gr_ratio'});

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
vehSpeed = d.vehVel_kmh(:);
totalDemand = d.torqueDemandTotal_Nm(:);
totalActual = d.torqueActualTotal_Nm(:);
posLimit = d.controllerTorquePositiveLimit_Nm(:);
negLimit = d.controllerTorqueNegativeLimit_Nm(:);
gearNum = d.gearNumber(:);
gearRatio = d.gearRatio(:);

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

driveShortfall = NaN(size(totalDemand));
regenShortfall = NaN(size(totalDemand));
validActual = isfinite(totalDemand) & isfinite(totalActual);
driveShortfall(validActual & driveMask) = max(totalDemand(validActual & driveMask) - totalActual(validActual & driveMask), 0);
regenShortfall(validActual & regenMask) = max(totalActual(validActual & regenMask) - totalDemand(validActual & regenMask), 0);

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

if any(validPosLimit)
    rows = RCA_AddKPI(rows, 'Near Positive Limit Share', 100 * RCA_FractionTrue(driveNearLimitMask, validPosLimit), '%', ...
        'Saturation', 'Power Train Controller', 'demand torque + max available torque', ...
        sprintf('Demand above %.0f%% of the positive limit is treated as near-limit drive operation.', config.Thresholds.LimitUsageFraction * 100));
    rows = RCA_AddKPI(rows, 'Mean Positive Torque Reserve', mean(driveReserve(validPosLimit), 'omitnan'), 'Nm', ...
        'Saturation', 'Power Train Controller', 'demand torque + max available torque', ...
        'Positive torque reserve is max available drive torque minus demanded drive torque.');
end

if any(validNegLimit)
    rows = RCA_AddKPI(rows, 'Near Regen Limit Share', 100 * RCA_FractionTrue(regenNearLimitMask, validNegLimit), '%', ...
        'Saturation', 'Power Train Controller', 'demand torque + min available torque', ...
        sprintf('Demand above %.0f%% of the recuperation limit magnitude is treated as near-limit regen operation.', config.Thresholds.LimitUsageFraction * 100));
    rows = RCA_AddKPI(rows, 'Mean Regen Torque Reserve', mean(regenReserve(validNegLimit), 'omitnan'), 'Nm', ...
        'Saturation', 'Power Train Controller', 'demand torque + min available torque', ...
        'Regen reserve is available recuperation torque magnitude minus demanded recuperation torque magnitude.');
end

if any(validActual)
    rows = RCA_AddKPI(rows, 'Total Torque Tracking MAE', mean(abs(totalDemand(validActual) - totalActual(validActual)), 'omitnan'), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Mean absolute total torque tracking error between commanded and delivered torque.');
    rows = RCA_AddKPI(rows, 'Drive Torque Shortfall 95th Percentile', RCA_Percentile(driveShortfall(isfinite(driveShortfall)), 95), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Positive torque under-delivery tail severity during propulsion.');
    rows = RCA_AddKPI(rows, 'Regen Torque Shortfall 95th Percentile', RCA_Percentile(regenShortfall(isfinite(regenShortfall)), 95), 'Nm', ...
        'Tracking', 'Power Train Controller', 'demand torque + actual torque', ...
        'Recuperation torque under-delivery tail severity during braking or lift-off regen.');
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

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'PowerTrainController');
plotFiles = localAppendPlotFile(plotFiles, localPlotCommandOverview(figureFolder, t, accPedal, brkPedal, totalDemand, totalActual, posLimit, negLimit, driveNearLimitMask, regenNearLimitMask, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotMotorSplit(figureFolder, t, em1Demand, em2Demand, totalDemand, accelPhase, brakePhase, cruisePhase, splitImbalanceDrive, splitImbalanceRegen, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotPedalResponse(figureFolder, accPedal, brkPedal, totalDemand, driveNearLimitMask, regenNearLimitMask, accelPhase, brakePhase, cruisePhase, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotGearContext(figureFolder, t, gearNum, gearRatio, totalDemand, posLimit, negLimit, shiftMask, config));
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
