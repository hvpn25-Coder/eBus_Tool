function integralValue = RCA_TrapzFinite(x, y)
% RCA_TrapzFinite  Integrate finite samples without bridging across NaN gaps.

x = double(x(:));
y = double(y(:));
n = min(numel(x), numel(y));
x = x(1:n);
y = y(1:n);

valid = isfinite(x) & isfinite(y);
if nnz(valid) < 2
    integralValue = 0;
    return;
end

validIdx = find(valid);
gapBreaks = [0; find(diff(validIdx) > 1); numel(validIdx)];
integralValue = 0;

for iSeg = 1:numel(gapBreaks) - 1
    segIdx = validIdx(gapBreaks(iSeg) + 1:gapBreaks(iSeg + 1));
    segX = x(segIdx);
    segY = y(segIdx);

    [segX, order] = sort(segX);
    segY = segY(order);
    [segXUnique, ~, groupIdx] = unique(segX);
    segYUnique = accumarray(groupIdx, segY, [], @mean);

    if numel(segXUnique) >= 2
        integralValue = integralValue + trapz(segXUnique, segYUnique);
    end
end
end
