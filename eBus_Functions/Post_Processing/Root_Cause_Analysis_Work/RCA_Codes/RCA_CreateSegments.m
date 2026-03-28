function segments = RCA_CreateSegments(derived, config)
% RCA_CreateSegments  Build meaningful trip segments from vehicle states.

t = localPrepareTimeVector(derived);
n = numel(t);
distanceStep = localPrepareVector(derived, 'distanceStep_km', n, 0);

if n == 0
    segments = localEmptySegmentsTable();
    return;
end

if n < 2
    segments = table(1, 1, 1, t(1), t(end), 0, sum(distanceStep, 'omitnan'), "Unknown", "Flat", "NormalAux", NaN, 0, ...
        'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
        'Duration_s', 'Distance_km', 'MotionClass', 'GradeClass', 'AuxClass', 'DominantGear', 'ShiftCount'});
    return;
end

vehVel = localPrepareVector(derived, 'vehVel_kmh', n, NaN);
vehAcc = localPrepareVector(derived, 'vehicleAcceleration_mps2', n, 0);
roadSlope = localPrepareVector(derived, 'roadSlope_pct', n, 0);
auxiliaryPower = localPrepareVector(derived, 'auxiliaryPower_kW', n, 0);
gear = localPrepareVector(derived, 'gearNumber', n, NaN);

motionClass = localMotionClass(vehVel, vehAcc, config);
gradeClass = repmat("Flat", n, 1);
gradeClass(roadSlope > config.Thresholds.UphillSlope_pct) = "Uphill";
gradeClass(roadSlope < config.Thresholds.DownhillSlope_pct) = "Downhill";

auxClass = repmat("NormalAux", n, 1);
auxClass(auxiliaryPower > config.Thresholds.HighAuxPower_kW) = "HighAux";

stateTag = motionClass + "|" + gradeClass + "|" + auxClass;
gearChange = [false; abs(diff(gear)) > 0 & ~isnan(diff(gear))];

boundary = [true; stateTag(2:end) ~= stateTag(1:end - 1)];
boundary = boundary | gearChange;
segmentStart = find(boundary);
if isempty(segmentStart)
    segmentStart = 1;
end
segmentEnd = [segmentStart(2:end) - 1; n];

[segmentStart, segmentEnd] = localMergeShortSegments(segmentStart, segmentEnd, t, config);
[segmentStart, segmentEnd] = localValidateSegments(segmentStart, segmentEnd, n);

if isempty(segmentStart) || isempty(segmentEnd)
    segmentStart = 1;
    segmentEnd = n;
end

segmentRows = cell(0, 12);
segmentCount = min(numel(segmentStart), numel(segmentEnd));
for iSeg = 1:segmentCount
    if segmentStart(iSeg) < 1 || segmentEnd(iSeg) > n || segmentStart(iSeg) > segmentEnd(iSeg)
        continue;
    end
    idx = segmentStart(iSeg):segmentEnd(iSeg);
    segmentRows(end + 1, :) = {iSeg, segmentStart(iSeg), segmentEnd(iSeg), ...
        t(segmentStart(iSeg)), t(segmentEnd(iSeg)), ...
        t(segmentEnd(iSeg)) - t(segmentStart(iSeg)), ...
        sum(distanceStep(idx), 'omitnan'), ...
        localModeValue(motionClass(idx)), localModeValue(gradeClass(idx)), localModeValue(auxClass(idx)), ...
        localDominantGear(gear(idx)), sum(gearChange(idx), 'omitnan')}; %#ok<AGROW>
end

if isempty(segmentRows)
    segments = table(1, 1, n, t(1), t(end), max(t(end) - t(1), 0), sum(distanceStep, 'omitnan'), ...
        localModeValue(motionClass), localModeValue(gradeClass), localModeValue(auxClass), ...
        localDominantGear(gear), sum(gearChange, 'omitnan'), ...
        'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
        'Duration_s', 'Distance_km', 'MotionClass', 'GradeClass', 'AuxClass', 'DominantGear', 'ShiftCount'});
else
    segments = cell2table(segmentRows, 'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', ...
        'StartTime_s', 'EndTime_s', 'Duration_s', 'Distance_km', 'MotionClass', ...
        'GradeClass', 'AuxClass', 'DominantGear', 'ShiftCount'});
end
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
segmentStart = segmentStart(:);
segmentEnd = segmentEnd(:);

if numel(segmentStart) ~= numel(segmentEnd)
    minLength = min(numel(segmentStart), numel(segmentEnd));
    segmentStart = segmentStart(1:minLength);
    segmentEnd = segmentEnd(1:minLength);
end

if numel(segmentStart) <= 1 || isempty(t)
    return;
end

keepMerging = true;
while keepMerging
    if numel(segmentStart) <= 1 || numel(segmentEnd) <= 1
        break;
    end

    keepMerging = false;
    durations = max(t(segmentEnd) - t(segmentStart), 0);
    shortIdx = find(durations < config.Thresholds.MinSegmentDuration_s, 1, 'first');
    if isempty(shortIdx)
        break;
    end

    if shortIdx == 1
        segmentStart(2) = segmentStart(1);
        segmentStart(1) = [];
        segmentEnd(1) = [];
    elseif shortIdx == numel(segmentStart)
        segmentEnd(shortIdx - 1) = segmentEnd(shortIdx);
        segmentStart(shortIdx) = [];
        segmentEnd(shortIdx) = [];
    else
        leftDuration = durations(shortIdx - 1);
        rightDuration = durations(shortIdx + 1);
        if leftDuration >= rightDuration
            segmentEnd(shortIdx - 1) = segmentEnd(shortIdx);
            segmentStart(shortIdx) = [];
            segmentEnd(shortIdx) = [];
        else
            segmentStart(shortIdx + 1) = segmentStart(shortIdx);
            segmentStart(shortIdx) = [];
            segmentEnd(shortIdx) = [];
        end
    end
    keepMerging = true;
end
end

function [segmentStart, segmentEnd] = localValidateSegments(segmentStart, segmentEnd, n)
segmentStart = segmentStart(:);
segmentEnd = segmentEnd(:);

if isempty(segmentStart) || isempty(segmentEnd)
    segmentStart = [];
    segmentEnd = [];
    return;
end

minLength = min(numel(segmentStart), numel(segmentEnd));
segmentPairs = [segmentStart(1:minLength), segmentEnd(1:minLength)];
segmentPairs = round(segmentPairs);
segmentPairs(:, 1) = max(1, min(n, segmentPairs(:, 1)));
segmentPairs(:, 2) = max(1, min(n, segmentPairs(:, 2)));
segmentPairs = sortrows(segmentPairs, 1);

validMask = segmentPairs(:, 1) <= segmentPairs(:, 2);
segmentPairs = segmentPairs(validMask, :);

if isempty(segmentPairs)
    segmentStart = [];
    segmentEnd = [];
    return;
end

segmentPairs(1, 1) = 1;
for iPair = 2:size(segmentPairs, 1)
    segmentPairs(iPair, 1) = max(segmentPairs(iPair, 1), segmentPairs(iPair - 1, 2) + 1);
end
segmentPairs(end, 2) = n;

validMask = segmentPairs(:, 1) <= segmentPairs(:, 2);
segmentPairs = segmentPairs(validMask, :);

segmentStart = segmentPairs(:, 1);
segmentEnd = segmentPairs(:, 2);
end

function t = localPrepareTimeVector(derived)
if isfield(derived, 'time_s') && ~isempty(derived.time_s)
    t = double(derived.time_s(:));
else
    t = zeros(0, 1);
end

if isempty(t)
    return;
end

isUsable = all(isfinite(t)) && all(diff(t) >= 0);
if isUsable
    return;
end

dt = diff(t);
dt = dt(isfinite(dt) & dt > 0);
if isempty(dt)
    sampleStep = 1;
else
    sampleStep = median(dt);
    if ~isfinite(sampleStep) || sampleStep <= 0
        sampleStep = 1;
    end
end

t = (0:sampleStep:sampleStep * (numel(t) - 1)).';
end

function vec = localPrepareVector(derived, fieldName, n, fillValue)
vec = repmat(fillValue, n, 1);
if n == 0 || ~isfield(derived, fieldName) || isempty(derived.(fieldName))
    return;
end

data = double(derived.(fieldName)(:));
if isscalar(data)
    vec = repmat(data, n, 1);
elseif numel(data) == n
    vec = data;
else
    sourceIndex = linspace(0, 1, numel(data));
    targetIndex = linspace(0, 1, n);
    vec = interp1(sourceIndex(:), data, targetIndex(:), 'linear', 'extrap');
end
end

function segments = localEmptySegmentsTable()
segments = table(zeros(0, 1), zeros(0, 1), zeros(0, 1), zeros(0, 1), zeros(0, 1), zeros(0, 1), zeros(0, 1), ...
    strings(0, 1), strings(0, 1), strings(0, 1), zeros(0, 1), zeros(0, 1), ...
    'VariableNames', {'SegmentID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
    'Duration_s', 'Distance_km', 'MotionClass', 'GradeClass', 'AuxClass', 'DominantGear', 'ShiftCount'});
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
