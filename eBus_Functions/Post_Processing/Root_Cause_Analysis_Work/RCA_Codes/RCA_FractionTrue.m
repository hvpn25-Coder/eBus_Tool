function fraction = RCA_FractionTrue(conditionMask, validMask)
% RCA_FractionTrue  Fraction of true samples over a valid-sample mask.

if nargin < 2 || isempty(validMask)
    validMask = true(size(conditionMask));
end

conditionMask = logical(conditionMask(:));
validMask = logical(validMask(:));
commonLength = min(numel(conditionMask), numel(validMask));
conditionMask = conditionMask(1:commonLength);
validMask = validMask(1:commonLength);

if ~any(validMask)
    fraction = NaN;
else
    fraction = sum(conditionMask(validMask)) / sum(validMask);
end
end
