function [segmentKPI, segmentSummary] = RCA_ComputeSegmentKPIs(derived, ~, ~, segments, config)
% RCA_ComputeSegmentKPIs  Compute segment-wise performance and efficiency KPIs.

t = derived.time_s;
if isempty(t) || isempty(segments) || height(segments) == 0
    segmentKPI = cell2table(cell(0, 8), 'VariableNames', {'SegmentID', 'KPIName', 'Value', 'Unit', ...
        'Category', 'Subsystem', 'SignalBasis', 'StatusNote'});
    segmentSummary = cell2table(cell(0, 40), 'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', ...
        'StartTime_s', 'EndTime_s', 'Duration_s', 'Distance_km', 'MotionClass', 'GradeClass', ...
        'AuxClass', 'DominantGear', 'ShiftCount', 'MeanSpeed_kmh', 'MeanSlope_pct', 'MeanSOC_pct', ...
        'MeanBatteryPower_kW', 'BatteryDischarge_kWh', 'BatteryRegen_kWh', 'Wh_per_km', ...
        'TrackingMAE_kmh', 'AuxEnergyShare_pct', 'LossEnergy_kWh', 'LossShare_pct', ...
        'MotorLossShare_pct', 'GearboxLossShare_pct', 'RollingLoadShare_pct', 'AeroLoadShare_pct', ...
        'ShiftRate_per_km', 'HuntingCount', 'BatteryLimitUse_pct', 'RegenRecovery_pct', ...
        'TorqueTrackingMAE_Nm', 'MotorHighSpeedShare_pct', 'EfficiencySeverity', 'PerformanceSeverity', ...
        'IsPoorEfficiency', 'IsPoorPerformance', 'IsHighLoss', 'PrimaryIssueTag', 'StatusNote'});
    return;
end

segmentRows = cell(0, 8);
summaryRows = cell(0, 40);

for iSeg = 1:height(segments)
    idx = segments.StartIndex(iSeg):segments.EndIndex(iSeg);
    segTime = t(idx);
    segDistance = sum(derived.distanceStep_km(idx), 'omitnan');
    battDischarge = RCA_TrapzFinite(segTime, max(derived.batteryPower_kW(idx), 0)) / 3600;
    battRegen = RCA_TrapzFinite(segTime, max(-derived.batteryPower_kW(idx), 0)) / 3600;
    auxEnergy = RCA_TrapzFinite(segTime, max(derived.auxiliaryPower_kW(idx), 0)) / 3600;
    battLoss = RCA_TrapzFinite(segTime, max(derived.batteryLossPower_kW(idx), 0)) / 3600;
    motorLoss = RCA_TrapzFinite(segTime, max(derived.motorLossPower_kW(idx), 0)) / 3600;
    gbxLoss = RCA_TrapzFinite(segTime, max(derived.gearboxLossPower_kW(idx), 0)) / 3600;
    fricEnergy = RCA_TrapzFinite(segTime, max(derived.frictionBrakePower_kW(idx), 0)) / 3600;
    lossEnergy = battLoss + motorLoss + gbxLoss + fricEnergy;
    if segDistance > config.General.MinimumDistanceForWhpkm_km
        whPerKm = 1000 * (battDischarge - battRegen) / segDistance;
    else
        whPerKm = NaN;
    end
    trackingMae = mean(abs(derived.speedError_kmh(idx)), 'omitnan');
    torqueTrackingMae = mean(abs(derived.torqueDemandTotal_Nm(idx) - derived.torqueActualTotal_Nm(idx)), 'omitnan');
    if battDischarge > 0
        auxShare = 100 * auxEnergy / battDischarge;
        lossShare = 100 * lossEnergy / battDischarge;
        motorLossShare = 100 * motorLoss / battDischarge;
        gbxLossShare = 100 * gbxLoss / battDischarge;
        rollingLoadShare = 100 * RCA_TrapzFinite(segTime, max(derived.rollingResistanceForce_N(idx), 0) .* max(derived.vehVel_mps(idx), 0) / 1000) / 3600 / battDischarge;
        aeroLoadShare = 100 * RCA_TrapzFinite(segTime, max(derived.aeroDragForce_N(idx), 0) .* max(derived.vehVel_mps(idx), 0) / 1000) / 3600 / battDischarge;
    else
        auxShare = NaN;
        lossShare = NaN;
        motorLossShare = NaN;
        gbxLossShare = NaN;
        rollingLoadShare = NaN;
        aeroLoadShare = NaN;
    end
    dischargePower = max(derived.batteryPower_kW(idx), 0);
    chargePower = max(-derived.batteryPower_kW(idx), 0);
    dischargeCurrent = max(derived.batteryCurrent_A(idx), 0);
    chargeCurrent = max(-derived.batteryCurrent_A(idx), 0);
    dischargePowerLimit = derived.battDischargePowerLimit_kW(idx);
    chargePowerLimit = derived.battChargePowerLimit_kW(idx);
    dischargeCurrentLimit = derived.battDischargeCurrentLimit_A(idx);
    chargeCurrentLimit = derived.battChargeCurrentLimit_A(idx);
    dischargePowerActive = dischargePower > 0 & isfinite(dischargePower) & isfinite(dischargePowerLimit) & dischargePowerLimit > 0;
    chargePowerActive = chargePower > 0 & isfinite(chargePower) & isfinite(chargePowerLimit) & chargePowerLimit > 0;
    dischargeCurrentActive = dischargeCurrent > 0 & isfinite(dischargeCurrent) & isfinite(dischargeCurrentLimit) & dischargeCurrentLimit > 0;
    chargeCurrentActive = chargeCurrent > 0 & isfinite(chargeCurrent) & isfinite(chargeCurrentLimit) & chargeCurrentLimit > 0;
    nearAnyLimit = false(size(dischargePower));
    nearAnyLimit(dischargePowerActive) = dischargePower(dischargePowerActive) > config.Thresholds.LimitUsageFraction .* dischargePowerLimit(dischargePowerActive);
    nearAnyLimit(chargePowerActive) = nearAnyLimit(chargePowerActive) | ...
        chargePower(chargePowerActive) > config.Thresholds.LimitUsageFraction .* chargePowerLimit(chargePowerActive);
    nearAnyLimit(dischargeCurrentActive) = nearAnyLimit(dischargeCurrentActive) | ...
        dischargeCurrent(dischargeCurrentActive) > config.Thresholds.LimitUsageFraction .* dischargeCurrentLimit(dischargeCurrentActive);
    nearAnyLimit(chargeCurrentActive) = nearAnyLimit(chargeCurrentActive) | ...
        chargeCurrent(chargeCurrentActive) > config.Thresholds.LimitUsageFraction .* chargeCurrentLimit(chargeCurrentActive);
    validLimitMask = dischargePowerActive | chargePowerActive | dischargeCurrentActive | chargeCurrentActive;
    if any(validLimitMask)
        batteryLimitUse = 100 * RCA_FractionTrue(nearAnyLimit, validLimitMask);
    else
        batteryLimitUse = NaN;
    end
    if (battRegen + fricEnergy) > 0
        regenRecovery = 100 * battRegen / (battRegen + fricEnergy);
    else
        regenRecovery = NaN;
    end
    meanSoc = mean(derived.batterySOC_pct(idx), 'omitnan');
    meanBatteryPower = mean(derived.batteryPower_kW(idx), 'omitnan');
    meanSpeed = mean(derived.vehVel_kmh(idx), 'omitnan');
    meanSlope = mean(derived.roadSlope_pct(idx), 'omitnan');
    motorSpeedAbs = abs(derived.motorSpeed_rpm(idx));
    validMotorSpeed = isfinite(motorSpeedAbs) & isfinite(derived.motorElectricalPower_kW(idx)) & abs(derived.motorElectricalPower_kW(idx)) > 1;
    tripMotorSpeedRef = max(abs(derived.motorSpeed_rpm), [], 'omitnan');
    if any(validMotorSpeed) && isfinite(tripMotorSpeedRef) && tripMotorSpeedRef > 0
        motorHighSpeedShare = 100 * RCA_FractionTrue( ...
            motorSpeedAbs > config.Thresholds.HighMotorEfficiencySpeedFraction * tripMotorSpeedRef, ...
            validMotorSpeed);
    else
        motorHighSpeedShare = NaN;
    end

    segGear = derived.gearNumber(idx);
    changeIdx = find(abs(diff(segGear)) > 0 & ~isnan(diff(segGear))) + 1;
    if segDistance > config.General.MinimumDistanceForWhpkm_km
        shiftRate = numel(changeIdx) / segDistance;
    else
        shiftRate = NaN;
    end
    huntingCount = 0;
    for iShift = 3:numel(changeIdx)
        if segGear(changeIdx(iShift)) == segGear(changeIdx(iShift - 2)) && ...
                (segTime(changeIdx(iShift)) - segTime(changeIdx(iShift - 2))) <= config.Thresholds.GearHuntingWindow_s
            huntingCount = huntingCount + 1;
        end
    end

    summaryRows(end + 1, :) = {segments.SegmentID(iSeg), segments.StartIndex(iSeg), segments.EndIndex(iSeg), ...
        segments.StartTime_s(iSeg), segments.EndTime_s(iSeg), segments.Duration_s(iSeg), segDistance, ...
        segments.MotionClass(iSeg), segments.GradeClass(iSeg), segments.AuxClass(iSeg), segments.DominantGear(iSeg), ...
        segments.ShiftCount(iSeg), meanSpeed, meanSlope, meanSoc, meanBatteryPower, battDischarge, battRegen, whPerKm, ...
        trackingMae, auxShare, lossEnergy, lossShare, motorLossShare, gbxLossShare, rollingLoadShare, aeroLoadShare, ...
        shiftRate, huntingCount, batteryLimitUse, regenRecovery, torqueTrackingMae, motorHighSpeedShare, ...
        NaN, NaN, false, false, false, "Unclassified", ""}; %#ok<AGROW>

    segmentRows = localAddSegmentKPI(segmentRows, segments.SegmentID(iSeg), 'Segment Distance', segDistance, 'km', 'Segment', 'Vehicle', 'veh_pos or integrated veh_vel', 'Segment distance.');
    segmentRows = localAddSegmentKPI(segmentRows, segments.SegmentID(iSeg), 'Segment Net Energy Intensity', whPerKm, 'Wh/km', 'Segment', 'Vehicle', 'batt_pwr + segment distance', 'Segment net electrical intensity using discharge-positive battery convention.');
    segmentRows = localAddSegmentKPI(segmentRows, segments.SegmentID(iSeg), 'Segment Tracking MAE', trackingMae, 'km/h', 'Segment', 'Vehicle', 'veh_des_vel + veh_vel', 'Segment tracking error.');
    segmentRows = localAddSegmentKPI(segmentRows, segments.SegmentID(iSeg), 'Segment Auxiliary Energy Share', auxShare, '%', 'Segment', 'Vehicle', 'auxiliary power + battery power', 'Auxiliary share relative to discharge-positive battery energy in this segment.');
    segmentRows = localAddSegmentKPI(segmentRows, segments.SegmentID(iSeg), 'Segment Loss Share', lossShare, '%', 'Segment', 'Vehicle', 'loss powers + battery power', 'Integrated loss share relative to discharge-positive battery energy in this segment.');
    segmentRows = localAddSegmentKPI(segmentRows, segments.SegmentID(iSeg), 'Segment Gear Shift Rate', shiftRate, 'shifts/km', 'Segment', 'Transmission', 'gr_num + segment distance', 'Gear change density inside the segment.');
end

segmentSummary = cell2table(summaryRows, 'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', ...
    'StartTime_s', 'EndTime_s', 'Duration_s', 'Distance_km', 'MotionClass', 'GradeClass', ...
    'AuxClass', 'DominantGear', 'ShiftCount', 'MeanSpeed_kmh', 'MeanSlope_pct', 'MeanSOC_pct', ...
    'MeanBatteryPower_kW', 'BatteryDischarge_kWh', 'BatteryRegen_kWh', 'Wh_per_km', ...
    'TrackingMAE_kmh', 'AuxEnergyShare_pct', 'LossEnergy_kWh', 'LossShare_pct', ...
    'MotorLossShare_pct', 'GearboxLossShare_pct', 'RollingLoadShare_pct', 'AeroLoadShare_pct', ...
    'ShiftRate_per_km', 'HuntingCount', 'BatteryLimitUse_pct', 'RegenRecovery_pct', ...
    'TorqueTrackingMAE_Nm', 'MotorHighSpeedShare_pct', 'EfficiencySeverity', 'PerformanceSeverity', ...
    'IsPoorEfficiency', 'IsPoorPerformance', 'IsHighLoss', 'PrimaryIssueTag', 'StatusNote'});

validEnergy = segmentSummary.Distance_km > config.General.MinimumDistanceForWhpkm_km & isfinite(segmentSummary.Wh_per_km);
tripMedianWhpkm = median(segmentSummary.Wh_per_km(validEnergy), 'omitnan');
if isnan(tripMedianWhpkm) || tripMedianWhpkm <= 0
    tripMedianWhpkm = max(median(segmentSummary.Wh_per_km, 'omitnan'), 1);
end

segmentSummary.EfficiencySeverity = segmentSummary.Wh_per_km ./ max(tripMedianWhpkm, eps);
torqueTrackingP75 = RCA_Percentile(segmentSummary.TorqueTrackingMAE_Nm, 75);
segmentSummary.PerformanceSeverity = max(segmentSummary.TrackingMAE_kmh ./ max(config.Thresholds.PoorTrackingError_kmh, eps), ...
    segmentSummary.TorqueTrackingMAE_Nm ./ max(torqueTrackingP75, eps));
segmentSummary.IsPoorEfficiency = segmentSummary.EfficiencySeverity > config.Thresholds.PoorEfficiencyMultiplier;
segmentSummary.IsPoorPerformance = segmentSummary.TrackingMAE_kmh > config.Thresholds.PoorTrackingError_kmh | ...
    segmentSummary.TorqueTrackingMAE_Nm > torqueTrackingP75;
segmentSummary.IsHighLoss = segmentSummary.LossShare_pct > config.Thresholds.HighLossShare_pct;
segmentSummary.PrimaryIssueTag = repmat("Mixed", height(segmentSummary), 1);
segmentSummary.PrimaryIssueTag(segmentSummary.IsPoorEfficiency & ~segmentSummary.IsPoorPerformance) = "Efficiency";
segmentSummary.PrimaryIssueTag(~segmentSummary.IsPoorEfficiency & segmentSummary.IsPoorPerformance) = "Performance";
segmentSummary.PrimaryIssueTag(segmentSummary.IsHighLoss) = "HighLoss";
segmentSummary.StatusNote = repmat("Segment summary computed from available signals with workbook sign conventions applied.", height(segmentSummary), 1);

segmentKPI = cell2table(segmentRows, 'VariableNames', {'SegmentID', 'KPIName', 'Value', 'Unit', ...
    'Category', 'Subsystem', 'SignalBasis', 'StatusNote'});
end

function rows = localAddSegmentKPI(rows, segmentID, kpiName, value, unit, category, subsystem, basis, note)
rows(end + 1, :) = {segmentID, string(kpiName), double(value), string(unit), string(category), ...
    string(subsystem), string(basis), string(note)};
end
