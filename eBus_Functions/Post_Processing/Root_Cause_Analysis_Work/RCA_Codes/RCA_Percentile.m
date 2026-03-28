function value = RCA_Percentile(data, pct)
% RCA_Percentile  Base-MATLAB percentile using linear interpolation.

data = double(data(:));
data = data(isfinite(data));
pct = double(pct);

if isempty(data) || isempty(pct)
    value = NaN(size(pct));
    return;
end

data = sort(data);
n = numel(data);
value = NaN(size(pct));

for iPct = 1:numel(pct)
    p = min(max(pct(iPct), 0), 100);
    if n == 1
        value(iPct) = data(1);
        continue;
    end

    pos = 1 + (n - 1) * p / 100;
    lo = floor(pos);
    hi = ceil(pos);
    if lo == hi
        value(iPct) = data(lo);
    else
        weight = pos - lo;
        value(iPct) = data(lo) * (1 - weight) + data(hi) * weight;
    end
end
end
