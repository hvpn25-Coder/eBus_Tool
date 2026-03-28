function segments = RCA_CreateSegments(derived, config)
% RCA_CreateSegments  Build meaningful trip segments from vehicle states.

t = derived.time_s;
n = numel(t);
if n < 2
    segments = table(1, 1, n, t(1), t(end), 0, 0, "Unknown", "Flat", "NormalAux", NaN, 0, ...
        'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
        'Duration_s', 'Distance_km', 'MotionClass', 'GradeClass', 'AuxClass', 'DominantGear', 'ShiftCount'});
    return;
end

motionClass = localMotionClass(derived.vehVel_kmh, derived.vehicleAcceleration_mps2, config);
gradeClass = repmat("Flat", n, 1);
gradeClass(derived.roadSlope_pct > config.Thresholds.UphillSlope_pct) = "Uphill";
gradeClass(derived.roadSlope_pct < config.Thresholds.DownhillSlope_pct) = "Downhill";

auxClass = repmat("NormalAux", n, 1);
auxClass(derived.auxiliaryPower_kW > config.Thresholds.HighAuxPower_kW) = "HighAux";

stateTag = motionClass + "|" + gradeClass + "|" + auxClass;
gear = derived.gearNumber;
gearChange = [false; abs(diff(gear)) > 0 & ~isnan(diff(gear))];

boundary = [true; stateTag(2:end) ~= stateTag(1:end - 1)];
boundary = boundary | gearChange;
segmentStart = find(boundary);
segmentEnd = [segmentStart(2:end) - 1; n];

[segmentStart, segmentEnd] = localMergeShortSegments(segmentStart, segmentEnd, t, config);

segmentRows = cell(0, 12);
for iSeg = 1:numel(segmentStart)
    idx = segmentStart(iSeg):segmentEnd(iSeg);
    segmentRows(end + 1, :) = {iSeg, segmentStart(iSeg), segmentEnd(iSeg), ...
        t(segmentStart(iSeg)), t(segmentEnd(iSeg)), ...
        t(segmentEnd(iSeg)) - t(segmentStart(iSeg)), ...
        sum(derived.distanceStep_km(idx), 'omitnan'), ...
        localModeValue(motionClass(idx)), localModeValue(gradeClass(idx)), localModeValue(auxClass(idx)), ...
        localDominantGear(gear(idx)), sum(gearChange(idx), 'omitnan')}; %#ok<AGROW>
end

segments = cell2table(segmentRows, 'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', ...
    'StartTime_s', 'EndTime_s', 'Duration_s', 'Distance_km', 'MotionClass', ...
    'GradeClass', 'AuxClass', 'DominantGear', 'ShiftCount'});
end

function motionClass = localMotionClass(speed, acceleration, config)
n = numel(speed);
motionClass = repmat("Cruise", n, 1);
motionClass(speed <= config.Thresholds.StopSpeed_kmh) = "Stop";
motionClass(speed > config.Thresholds.StopSpeed_kmh & speed <= config.Thresholds.CreepSpeed_kmh) = "Creep";
motionClass(speed > config.Thresholds.CreepSpeed_kmh & acceleration >= config.Thresholds.SignificantAccel_mps2) = "Accelerate";
motionClass(speed > config.Thresholds.CreepSpeed_kmh & acceleration <= config.Thresholds.SignificantDecel_mps2) = "Brake";
motionClass(speed > config.Thresholds.CreepSpeed_kmh & abs(acceleration) <= config.Thresholds.CruiseAccelAbs_mps2) = "Cruise";
end

function [segmentStart, segmentEnd] = localMergeShortSegments(segmentStart, segmentEnd, t, config)
if numel(segmentStart) <= 1
    return;
end

keepMerging = true;
while keepMerging
    keepMerging = false;
    durations = t(segmentEnd) - t(segmentStart);
    shortIdx = find(durations < config.Thresholds.MinSegmentDuration_s, 1, 'first');
    if isempty(shortIdx)
        break;
    end

    if shortIdx == 1
        segmentStart(2) = segmentStart(1);
        segmentStart(1) = [];
        segmentEnd(1) = [];
    else
        segmentEnd(shortIdx - 1) = segmentEnd(shortIdx);
        segmentStart(shortIdx) = [];
        segmentEnd(shortIdx) = [];
    end
    keepMerging = true;
end
end

function modeValue = localModeValue(values)
values = values(strlength(values) > 0);
if isempty(values)
    modeValue = "Unknown";
    return;
end
uniqueValues = unique(values);
counts = zeros(size(uniqueValues));
for iValue = 1:numel(uniqueValues)
    counts(iValue) = sum(values == uniqueValues(iValue));
end
[~, idx] = max(counts);
modeValue = uniqueValues(idx);
end

function dominantGear = localDominantGear(gear)
gear = gear(~isnan(gear));
if isempty(gear)
    dominantGear = NaN;
    return;
end
uniqueGear = unique(gear);
counts = zeros(size(uniqueGear));
for iGear = 1:numel(uniqueGear)
    counts(iGear) = sum(gear == uniqueGear(iGear));
end
[~, idx] = max(counts);
dominantGear = uniqueGear(idx);
end
