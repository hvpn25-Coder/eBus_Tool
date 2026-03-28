function [signalStore, referenceInfo] = RCA_AlignSignalStore(signalStore, rawData, config, timeSignalNames)
% RCA_AlignSignalStore  Align all dynamic signals to a common reference time.

if nargin < 3 || isempty(config)
    config = RCA_Config();
end
if nargin < 4 || isempty(timeSignalNames)
    timeSignalNames = strings(0, 1);
else
    timeSignalNames = string(timeSignalNames(:));
end

[referenceTime, referenceSource, referenceMethod] = localSelectReferenceTime(signalStore, rawData, config, timeSignalNames);
[referenceTime, timeWasSanitized] = localSanitizeReferenceTime(referenceTime);
if timeWasSanitized
    referenceMethod = string(referenceMethod) + "|Sanitized";
end
nRef = numel(referenceTime);

signalFields = fieldnames(signalStore);
for iField = 1:numel(signalFields)
    signal = signalStore.(signalFields{iField});
    if ~isfield(signal, 'Available') || ~signal.Available
        continue;
    end

    data = signal.Data;
    time = signal.Time;
    interpolationMethod = localInterpolationMethod(signal, config);
    if isempty(data)
        signal.AlignedData = [];
        signal.AlignedTime = [];
        signalStore.(signalFields{iField}) = signal;
        continue;
    end

    if isscalar(data)
        signal.AlignedData = repmat(double(data), nRef, 1);
        signal.AlignedTime = referenceTime;
        signal.Note = string(signal.Note) + " | Scalar expanded to reference length.";
    elseif isempty(time)
        if numel(data) == nRef
            signal.AlignedData = double(data(:));
            signal.AlignedTime = referenceTime;
        else
            sourceIndex = linspace(referenceTime(1), referenceTime(end), numel(data));
            signal.AlignedData = interp1(sourceIndex(:), double(data(:)), referenceTime, ...
                interpolationMethod, 'extrap');
            signal.AlignedTime = referenceTime;
            signal.Note = string(signal.Note) + " | Sample-index interpolation used because time was missing.";
            signal.Approximate = true;
            signal.Confidence = "Medium";
        end
    else
        time = double(time(:));
        data = double(data(:));
        if numel(time) ~= numel(data)
            commonLength = min(numel(time), numel(data));
            time = time(1:commonLength);
            data = data(1:commonLength);
            signal.Note = string(signal.Note) + " | Time and data length mismatch trimmed to common length.";
        end

        [time, data, prepNote] = localPrepareInterpolationSeries(time, data);
        if strlength(prepNote) > 0
            signal.Note = string(signal.Note) + " | " + prepNote;
        end

        if isempty(time)
            if numel(data) == nRef
                signal.AlignedData = data(:);
                signal.AlignedTime = referenceTime;
                signal.Note = string(signal.Note) + " | Original time basis was unusable, so aligned data was accepted by sample count.";
                signal.Approximate = true;
                signal.Confidence = "Medium";
            elseif numel(data) == 1
                signal.AlignedData = repmat(double(data(1)), nRef, 1);
                signal.AlignedTime = referenceTime;
                signal.Note = string(signal.Note) + " | Original time basis was unusable, so scalar expansion was used.";
                signal.Approximate = true;
                signal.Confidence = "Medium";
            else
                sourceIndex = linspace(referenceTime(1), referenceTime(end), numel(data));
                signal.AlignedData = interp1(sourceIndex(:), double(data(:)), referenceTime, interpolationMethod, 'extrap');
                signal.AlignedTime = referenceTime;
                signal.Note = string(signal.Note) + " | Original time basis was unusable, so sample-index interpolation was used.";
                signal.Approximate = true;
                signal.Confidence = "Medium";
            end
        elseif numel(time) == 1
            signal.AlignedData = repmat(double(data(1)), nRef, 1);
            signal.AlignedTime = referenceTime;
            signal.Note = string(signal.Note) + " | Single unique timestamp remained after time cleanup; scalar expansion used.";
            signal.Approximate = true;
            signal.Confidence = "Medium";
        elseif numel(time) == nRef && max(abs(time - referenceTime), [], 'omitnan') < 1e-6
            signal.AlignedData = data(:);
            signal.AlignedTime = referenceTime;
        else
            signal.AlignedData = interp1(time, data, referenceTime, interpolationMethod, 'extrap');
            signal.AlignedTime = referenceTime;
        end
    end

    signalStore.(signalFields{iField}) = signal;
end

referenceInfo = struct('Time_s', referenceTime, 'Source', string(referenceSource), 'Method', string(referenceMethod));
end

function [referenceTime, referenceSource, referenceMethod] = localSelectReferenceTime(signalStore, rawData, config, timeSignalNames)
referenceTime = [];
referenceSource = "";
referenceMethod = "SignalReference";

preferredSignals = unique([timeSignalNames; string(config.PreferredReferenceSignals(:))], 'stable');
for iPreferred = 1:numel(preferredSignals)
    candidate = RCA_GetSignalData(signalStore, preferredSignals(iPreferred));
    [candidateTime, candidateMethod] = localResolveSignalTime(candidate, config, timeSignalNames);
    if ~isempty(candidateTime)
        referenceTime = candidateTime;
        referenceSource = candidate.Name;
        referenceMethod = candidateMethod;
        return;
    end
end

signalFields = fieldnames(signalStore);
for iField = 1:numel(signalFields)
    signal = signalStore.(signalFields{iField});
    [candidateTime, candidateMethod] = localResolveSignalTime(signal, config, timeSignalNames);
    if ~isempty(candidateTime)
        referenceTime = candidateTime;
        referenceSource = signal.Name;
        referenceMethod = candidateMethod;
        return;
    end
end

if ~isempty(rawData.DefaultTime)
    referenceTime = double(rawData.DefaultTime(:));
    referenceSource = rawData.DefaultTimeSource;
    referenceMethod = "DatasetDefaultTime";
    return;
end

maxLength = 0;
for iField = 1:numel(signalFields)
    signal = signalStore.(signalFields{iField});
    if signal.Available && isnumeric(signal.Data) && numel(signal.Data) > maxLength
        maxLength = numel(signal.Data);
        referenceSource = signal.Name;
    end
end

if maxLength >= config.General.MinimumDynamicSamples
    referenceTime = (0:maxLength - 1)';
    referenceMethod = "SampleIndex";
else
    referenceTime = (0:1)';
    referenceMethod = "FallbackSampleIndex";
end
end

function method = localInterpolationMethod(signal, config)
signalText = lower(strjoin(string({signal.Name, signal.Description}), ' '));
if any(contains(signalText, string(config.SignalFallback.DiscreteTokenList)))
    method = 'nearest';
else
    method = 'linear';
end
end

function [candidateTime, candidateMethod] = localResolveSignalTime(signal, config, timeSignalNames)
candidateTime = [];
candidateMethod = "SignalReference";
if ~isfield(signal, 'Available') || ~signal.Available
    return;
end

if ~isempty(signal.Time) && numel(signal.Time) >= config.General.MinimumDynamicSamples && localIsUsableTimeVector(signal.Time)
    candidateTime = double(signal.Time(:));
    candidateMethod = "SignalReference";
    return;
end

if localIsExplicitTimeSignal(signal, timeSignalNames) && ~isempty(signal.Data) && ...
        numel(signal.Data) >= config.General.MinimumDynamicSamples && localIsUsableTimeVector(signal.Data)
    candidateTime = double(signal.Data(:));
    candidateMethod = "WorkbookTimeSignal";
end
end

function tf = localIsExplicitTimeSignal(signal, timeSignalNames)
signalNameKey = localNormalizeText(signal.Name);
descriptionKey = localNormalizeText(signal.Description);
unitKey = localNormalizeText(signal.Unit);
timeSignalNameKeys = localNormalizeText(timeSignalNames);

nameMatchesWorkbook = any(strcmp(signalNameKey, timeSignalNameKeys));
hasTimeText = strcmp(descriptionKey, "time") || strcmp(signalNameKey, "time") || ...
    strcmp(signalNameKey, "timesim") || strcmp(signalNameKey, "simtime") || ...
    contains(descriptionKey, "time") || contains(signalNameKey, "time");
hasSecondUnit = any(strcmp(unitKey, ["s", "sec", "secs", "second", "seconds"]));
tf = nameMatchesWorkbook || (hasTimeText && hasSecondUnit);
end

function tf = localIsUsableTimeVector(value)
value = double(value(:));
tf = numel(value) >= 3 && all(isfinite(value)) && all(diff(value) >= 0);
end

function normalized = localNormalizeText(textValue)
textValue = string(textValue);
normalized = strings(size(textValue));
for iValue = 1:numel(textValue)
    normalized(iValue) = lower(regexprep(textValue(iValue), '[^a-zA-Z0-9]', ''));
end
end

function [timeVector, wasSanitized] = localSanitizeReferenceTime(timeVector)
wasSanitized = false;
timeVector = double(timeVector(:));
originalLength = numel(timeVector);

if isempty(timeVector)
    timeVector = (0:1).';
    wasSanitized = true;
    return;
end

[timeVector, wasSanitized, requiresFallback] = localSanitizeTimeVector(timeVector);
if requiresFallback
    positiveDt = diff(timeVector);
    positiveDt = positiveDt(isfinite(positiveDt) & positiveDt > 0);
    if isempty(positiveDt)
        sampleStep = 1;
    else
        sampleStep = median(positiveDt);
        if ~isfinite(sampleStep) || sampleStep <= 0
            sampleStep = 1;
        end
    end
    fallbackLength = max(originalLength, 2);
    timeVector = (0:sampleStep:sampleStep * (fallbackLength - 1)).';
    wasSanitized = true;
end
end

function [time, data, note] = localPrepareInterpolationSeries(time, data)
note = "";
time = double(time(:));
data = double(data(:));

if isempty(time) || isempty(data)
    time = [];
    data = [];
    note = "Signal time or data was empty after cleanup.";
    return;
end

validMask = isfinite(time) & isfinite(data);
if ~all(validMask)
    time = time(validMask);
    data = data(validMask);
    note = localAppendNote(note, "Invalid time/data samples removed before alignment.");
end

if isempty(time) || isempty(data)
    time = [];
    data = [];
    note = localAppendNote(note, "No finite time-data samples remained.");
    return;
end

    wasSorted = false;
    if any(diff(time) < 0)
        [time, sortIdx] = sort(time);
        data = data(sortIdx);
        wasSorted = true;
    end

    originalLength = numel(time);
    [time, uniqueIdx] = unique(time, 'last');
    data = data(uniqueIdx);

    if wasSorted || numel(time) < originalLength
        note = localAppendNote(note, "Signal time vector was sorted and duplicate timestamps were collapsed.");
    end

    if numel(time) < 2
        time = [];
        note = localAppendNote(note, "Signal time vector could not provide a unique increasing basis.");
        return;
    end
end

function [timeVector, wasSanitized, requiresFallback] = localSanitizeTimeVector(timeVector)
timeVector = double(timeVector(:));
wasSanitized = false;
requiresFallback = false;

if isempty(timeVector)
    requiresFallback = true;
    return;
end

finiteMask = isfinite(timeVector);
if ~all(finiteMask)
    timeVector = timeVector(finiteMask);
    wasSanitized = true;
end

if isempty(timeVector)
    requiresFallback = true;
    return;
end

if all(diff(timeVector) > 0)
    return;
end

if all(diff(timeVector) >= 0)
    originalLength = numel(timeVector);
    timeVector = unique(timeVector, 'last');
    wasSanitized = wasSanitized || (numel(timeVector) < originalLength);
    if numel(timeVector) < 2
        requiresFallback = true;
    end
    return;
end

[timeVector, sortIdx] = sort(timeVector);
wasSanitized = wasSanitized || any(sortIdx(:) ~= (1:numel(sortIdx)).');
originalLength = numel(timeVector);
timeVector = unique(timeVector, 'last');
wasSanitized = wasSanitized || (numel(timeVector) < originalLength);
if numel(timeVector) < 2
    requiresFallback = true;
end
end

function note = localAppendNote(note, fragment)
fragment = string(fragment);
if strlength(note) == 0
    note = fragment;
else
    note = note + " " + fragment;
end
end
