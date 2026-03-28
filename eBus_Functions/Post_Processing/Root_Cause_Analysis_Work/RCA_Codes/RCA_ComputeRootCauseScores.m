function [rootCauseRanking, badSegmentTable, narrative, optimizationTable] = RCA_ComputeRootCauseScores(~, ~, ~, segmentSummary, config)
% RCA_ComputeRootCauseScores  Rank likely causal drivers for poor segments.

if isempty(segmentSummary) || height(segmentSummary) == 0
    rootCauseRanking = table(zeros(0, 1), zeros(0, 1), zeros(0, 1), zeros(0, 1), strings(0, 1), zeros(0, 1), zeros(0, 1), strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'SegmentID', 'StartTime_s', 'EndTime_s', 'CauseRank', 'CauseName', 'Score', 'Contribution_pct', 'EvidenceSignals', 'Confidence', 'Narrative'});
    badSegmentTable = table(zeros(0, 1), zeros(0, 1), zeros(0, 1), strings(0, 1), strings(0, 1), zeros(0, 1), strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'SegmentID', 'StartTime_s', 'EndTime_s', 'IssueType', 'PrimaryCause', 'PrimaryContribution_pct', 'Confidence', 'EvidenceSignals', 'Narrative'});
    narrative = "No segments were available for RCA scoring.";
    optimizationTable = table(strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'Subsystem', 'Recommendation', 'Evidence'});
    return;
end

candidateMask = segmentSummary.IsPoorEfficiency | segmentSummary.IsPoorPerformance | segmentSummary.IsHighLoss;
if ~any(candidateMask)
    [~, order] = sort(max(segmentSummary.EfficiencySeverity, segmentSummary.PerformanceSeverity), 'descend');
    candidateMask(order(1:min(config.Thresholds.WorstSegmentCount, numel(order)))) = true;
end

badSegments = segmentSummary(candidateMask, :);
rankingRows = cell(0, 10);
summaryRows = cell(0, 9);
narrative = strings(0, 1);
torqueTrackingP75 = max(RCA_Percentile(segmentSummary.TorqueTrackingMAE_Nm, 75), 1);

for iSeg = 1:height(badSegments)
    seg = badSegments(iSeg, :);
    factorNames = {'SlopeLoad', 'AuxiliaryLoad', 'LowSOC', 'BatteryLimit', 'GearOperation', ...
        'RollingResistance', 'AeroDrag', 'ControllerTracking', 'MotorInverter', 'TransmissionLoss', 'RegenMiss'};
    factorScores = [ ...
        config.RootCauseWeights.Slope * localNormalize(max(seg.MeanSlope_pct, 0), config.Thresholds.SteepSlope_pct), ...
        config.RootCauseWeights.Auxiliary * localNormalize(seg.AuxEnergyShare_pct, config.Thresholds.HighAuxShare_pct), ...
        config.RootCauseWeights.LowSOC * localNormalize(max(config.Thresholds.LowSOC_pct - seg.MeanSOC_pct, 0), max(config.Thresholds.LowSOC_pct - config.Thresholds.CriticalSOC_pct, 1)), ...
        config.RootCauseWeights.BatteryLimit * localNormalize(seg.BatteryLimitUse_pct, 25), ...
        config.RootCauseWeights.GearOperation * min(2, 0.6 * localNormalize(seg.ShiftRate_per_km, config.Thresholds.GearShiftRate_perkm) + ...
            0.8 * localNormalize(seg.HuntingCount, 1) + 0.5 * localNormalize(seg.MotorHighSpeedShare_pct, 20)), ...
        config.RootCauseWeights.RollingResistance * localNormalize(seg.RollingLoadShare_pct, 15), ...
        config.RootCauseWeights.AeroDrag * localNormalize(seg.AeroLoadShare_pct, 15), ...
        config.RootCauseWeights.ControllerTracking * max(localNormalize(seg.TrackingMAE_kmh, config.Thresholds.PoorTrackingError_kmh), ...
            localNormalize(seg.TorqueTrackingMAE_Nm, torqueTrackingP75)), ...
        config.RootCauseWeights.MotorInverter * localNormalize(seg.MotorLossShare_pct, config.Thresholds.HighLossShare_pct), ...
        config.RootCauseWeights.Transmission * localNormalize(seg.GearboxLossShare_pct, config.Thresholds.HighLossShare_pct), ...
        config.RootCauseWeights.RegenMiss * localNormalize(max(100 - seg.RegenRecovery_pct, 0), 100 * (1 - config.Thresholds.PoorRegenRecoveryFraction))];

    totalScore = sum(factorScores);
    if totalScore <= 0
        factorScores = ones(size(factorScores));
        totalScore = sum(factorScores);
    end
    contributionPct = 100 * factorScores / totalScore;
    [~, order] = sort(factorScores, 'descend');

    confidence = localConfidence(sum(factorScores > 0));
    topNarrative = sprintf('Segment %d (%.1f s to %.1f s) shows %.1f Wh/km and %.1f km/h tracking MAE. Leading drivers: %s %.1f%%, %s %.1f%%, %s %.1f%%.', ...
        seg.SegmentID, seg.StartTime_s, seg.EndTime_s, seg.Wh_per_km, seg.TrackingMAE_kmh, ...
        factorNames{order(1)}, contributionPct(order(1)), factorNames{order(2)}, contributionPct(order(2)), factorNames{order(3)}, contributionPct(order(3)));
    narrative(end + 1) = string(topNarrative);

    for iRank = 1:numel(order)
        rankingRows(end + 1, :) = {seg.SegmentID, seg.StartTime_s, seg.EndTime_s, iRank, ...
            string(factorNames{order(iRank)}), factorScores(order(iRank)), contributionPct(order(iRank)), ...
            string(localEvidenceSignals(factorNames{order(iRank)})), string(confidence), string(topNarrative)}; %#ok<AGROW>
    end

    summaryRows(end + 1, :) = {seg.SegmentID, seg.StartTime_s, seg.EndTime_s, seg.PrimaryIssueTag, ...
        string(factorNames{order(1)}), contributionPct(order(1)), string(confidence), ...
        string(localEvidenceSignals(factorNames{order(1)})), string(topNarrative)}; %#ok<AGROW>
end

rootCauseRanking = cell2table(rankingRows, 'VariableNames', {'SegmentID', 'StartTime_s', 'EndTime_s', ...
    'CauseRank', 'CauseName', 'Score', 'Contribution_pct', 'EvidenceSignals', 'Confidence', 'Narrative'});
badSegmentTable = cell2table(summaryRows, 'VariableNames', {'SegmentID', 'StartTime_s', 'EndTime_s', ...
    'IssueType', 'PrimaryCause', 'PrimaryContribution_pct', 'Confidence', 'EvidenceSignals', 'Narrative'});

topCauses = unique(badSegmentTable.PrimaryCause, 'stable');
optRows = cell(0, 3);
for iCause = 1:numel(topCauses)
    [recommendation, evidence] = localRecommendation(topCauses(iCause));
    optRows(end + 1, :) = {localSubsystemOwner(topCauses(iCause)), recommendation, evidence}; %#ok<AGROW>
end
optimizationTable = cell2table(optRows, 'VariableNames', {'Subsystem', 'Recommendation', 'Evidence'});
end

function normalized = localNormalize(value, threshold)
if isnan(value) || threshold <= 0
    normalized = 0;
else
    normalized = max(0, value) / threshold;
end
end

function confidence = localConfidence(activeFactors)
if activeFactors >= 6
    confidence = "High";
elseif activeFactors >= 3
    confidence = "Medium";
else
    confidence = "Low";
end
end

function evidence = localEvidenceSignals(causeName)
switch char(causeName)
    case 'SlopeLoad'
        evidence = 'road_slp, grad_force, veh_vel';
    case 'AuxiliaryLoad'
        evidence = 'aux_curr, aux_volt, batt_pwr';
    case 'LowSOC'
        evidence = 'batt_soc';
    case 'BatteryLimit'
        evidence = 'batt_pwr, batt_curr, BMS limits';
    case 'GearOperation'
        evidence = 'gr_num, gr_ratio, motor speed, gearbox loss';
    case 'RollingResistance'
        evidence = 'roll_res_force, veh_vel';
    case 'AeroDrag'
        evidence = 'aero_drag_force, veh_vel';
    case 'ControllerTracking'
        evidence = 'veh_des_vel, veh_vel, torque demand, torque actual';
    case 'MotorInverter'
        evidence = 'motor power, motor loss power';
    case 'TransmissionLoss'
        evidence = 'gbx_pwr_loss, gr_num';
    case 'RegenMiss'
        evidence = 'fric_brk_pwr, batt_pwr';
    otherwise
        evidence = 'Available segment signals';
end
end

function owner = localSubsystemOwner(causeName)
switch char(causeName)
    case {'SlopeLoad', 'RollingResistance', 'AeroDrag'}
        owner = "Vehicle Dynamics";
    case {'AuxiliaryLoad'}
        owner = "Auxiliary Load";
    case {'LowSOC'}
        owner = "Battery";
    case {'BatteryLimit'}
        owner = "Battery Management System";
    case {'GearOperation', 'TransmissionLoss'}
        owner = "Transmission";
    case {'ControllerTracking'}
        owner = "Power Train Controller";
    case {'MotorInverter'}
        owner = "Electric Drive";
    case {'RegenMiss'}
        owner = "Pneumatic Brake System";
    otherwise
        owner = "Vehicle";
end
end

function [recommendation, evidence] = localRecommendation(causeName)
switch char(causeName)
    case 'SlopeLoad'
        recommendation = "Normalize route comparisons for grade severity and add grade-compensation logic review where uphill segments dominate.";
        evidence = "Multiple bad segments are grade-driven.";
    case 'AuxiliaryLoad'
        recommendation = "Reduce continuous auxiliary demand and review idle/peak accessory control logic.";
        evidence = "Bad segments show high auxiliary energy share.";
    case 'LowSOC'
        recommendation = "Review usable SoC window and low-SoC control behaviour before blaming propulsion hardware for weak performance.";
        evidence = "Low-SoC operation repeatedly coincides with poor segments.";
    case 'BatteryLimit'
        recommendation = "Revisit charge/discharge limit calibration and battery capability assumptions to reduce active limiting.";
        evidence = "Battery power/current frequently run near configured limits.";
    case 'GearOperation'
        recommendation = "Retune shift schedule and hysteresis to reduce hunting and keep motor operation inside a more efficient region.";
        evidence = "Poor segments show high shift density or gear instability.";
    case 'RollingResistance'
        recommendation = "Recheck tyre rolling-loss assumptions and wheel-loss modelling against expected vehicle configuration.";
        evidence = "Rolling resistance is a repeated efficiency driver.";
    case 'AeroDrag'
        recommendation = "Review aero drag assumptions and the effect of high-speed operation on range-sensitive segments.";
        evidence = "Aerodynamic load repeatedly appears in poor segments.";
    case 'ControllerTracking'
        recommendation = "Tighten torque/speed tracking logic and separate supervisory limits from response delays.";
        evidence = "Tracking shortfall is a repeated poor-performance driver.";
    case 'MotorInverter'
        recommendation = "Shift operating points toward a better motor/inverter efficiency region or recalibrate current split and torque scheduling.";
        evidence = "Motor loss share is repeatedly high.";
    case 'TransmissionLoss'
        recommendation = "Review gearbox loss model and ratio usage to cut transmission-specific energy loss.";
        evidence = "Transmission loss share is repeatedly high.";
    case 'RegenMiss'
        recommendation = "Improve regen blending and charge-acceptance coordination so braking opportunity is converted to recovered energy.";
        evidence = "Recovered braking fraction is repeatedly poor.";
    otherwise
        recommendation = "Inspect segment evidence manually and add missing logging for stronger causality.";
        evidence = "Causal confidence is limited.";
end
end
