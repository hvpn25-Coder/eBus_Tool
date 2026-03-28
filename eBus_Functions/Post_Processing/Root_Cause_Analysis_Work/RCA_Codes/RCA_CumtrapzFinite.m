function cumulative = RCA_CumtrapzFinite(x, y)
% RCA_CumtrapzFinite  Cumulative trapezoidal integral with NaN-gap protection.

originalSize = size(y);
x = double(x(:));
y = double(y(:));
n = min(numel(x), numel(y));
x = x(1:n);
y = y(1:n);

cumulative = NaN(n, 1);
valid = isfinite(x) & isfinite(y);
if ~any(valid)
    cumulative = reshape(cumulative, originalSize);
    return;
end

runningValue = 0;
previousValid = false;
prevX = NaN;
prevY = NaN;

for iPoint = 1:n
    if valid(iPoint)
        if previousValid
            dx = x(iPoint) - prevX;
            if dx >= 0
                runningValue = runningValue + 0.5 * (prevY + y(iPoint)) * dx;
            end
        end
        prevX = x(iPoint);
        prevY = y(iPoint);
        previousValid = true;
    else
        previousValid = false;
    end
    cumulative(iPoint) = runningValue;
end

cumulative = reshape(cumulative, originalSize);
end
