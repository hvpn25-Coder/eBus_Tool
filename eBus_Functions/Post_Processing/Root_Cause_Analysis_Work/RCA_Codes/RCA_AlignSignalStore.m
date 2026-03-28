function [signalStore, referenceInfo] = RCA_AlignSignalStore(signalStore, rawData, config)
% RCA_AlignSignalStore  Align all dynamic signals to a common reference time.

if nargin < 3 || isempty(config)
    config = RCA_Config();
end

[referenceTime, referenceSource, referenceMethod] = localSelectReferenceTime(signalStore, rawData, config);
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
                localInterpolationMethod(signal, config), 'extrap');
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

        if numel(time) == nRef && max(abs(time - referenceTime), [], 'omitnan') < 1e-6
            signal.AlignedData = data(:);
            signal.AlignedTime = referenceTime;
        else
            signal.AlignedData = interp1(time, data, referenceTime, localInterpolationMethod(signal, config), 'extrap');
            signal.AlignedTime = referenceTime;
        end
    end

    signalStore.(signalFields{iField}) = signal;
end

referenceInfo = struct('Time_s', referenceTime, 'Source', string(referenceSource), 'Method', string(referenceMethod));
end

function [referenceTime, referenceSource, referenceMethod] = localSelectReferenceTime(signalStore, rawData, config)
referenceTime = [];
referenceSource = "";
referenceMethod = "SignalReference";

for iPreferred = 1:numel(config.PreferredReferenceSignals)
    candidate = RCA_GetSignalData(signalStore, config.PreferredReferenceSignals{iPreferred});
    if candidate.Available && ~isempty(candidate.Time) && numel(candidate.Time) >= config.General.MinimumDynamicSamples
        referenceTime = double(candidate.Time(:));
        referenceSource = candidate.Name;
        return;
    end
end

signalFields = fieldnames(signalStore);
for iField = 1:numel(signalFields)
    signal = signalStore.(signalFields{iField});
    if signal.Available && ~isempty(signal.Time) && numel(signal.Time) >= config.General.MinimumDynamicSamples
        referenceTime = double(signal.Time(:));
        referenceSource = signal.Name;
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
signalText = lower(strjoin([signal.Name, signal.Description], ' '));
if any(contains(signalText, string(config.SignalFallback.DiscreteTokenList)))
    method = 'nearest';
else
    method = 'linear';
end
end

function [timeVector, wasSanitized] = localSanitizeReferenceTime(timeVector)
wasSanitized = false;
timeVector = double(timeVector(:));

if isempty(timeVector)
    timeVector = (0:1).';
    wasSanitized = true;
    return;
end

isUsable = all(isfinite(timeVector)) && all(diff(timeVector) >= 0);
if isUsable
    return;
end

dt = diff(timeVector);
dt = dt(isfinite(dt) & dt > 0);
if isempty(dt)
    sampleStep = 1;
else
    sampleStep = median(dt);
    if ~isfinite(sampleStep) || sampleStep <= 0
        sampleStep = 1;
    end
end

timeVector = (0:sampleStep:sampleStep * (numel(timeVector) - 1)).';
wasSanitized = true;
end
