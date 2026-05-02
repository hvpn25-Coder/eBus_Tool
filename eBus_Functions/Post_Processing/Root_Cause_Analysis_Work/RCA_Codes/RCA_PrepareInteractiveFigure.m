function RCA_PrepareInteractiveFigure(figHandle)
% RCA_PrepareInteractiveFigure  Enable standard MATLAB interaction on RCA figures.
%
% RCA figures are saved for both Word embedding and MATLAB review. This
% helper keeps the regular MATLAB figure toolbar available and links subplot
% x-axes when they clearly share the same x-domain, which is typically the
% common time base used in RCA dashboards.

if nargin < 1 || isempty(figHandle) || ~isgraphics(figHandle)
    return;
end

try
    set(figHandle, 'MenuBar', 'figure', 'Toolbar', 'figure');
catch
end

axesHandles = localUsableAxes(figHandle);
for iAxis = 1:numel(axesHandles)
    try
        enableDefaultInteractivity(axesHandles(iAxis));
    catch
    end
    try
        if exist('axtoolbar', 'file') == 2
            axtoolbar(axesHandles(iAxis), {'export', 'datacursor', 'pan', 'zoomin', 'zoomout', 'restoreview'});
        end
    catch
    end
end

localLinkCompatibleXAxis(axesHandles);
end

function axesHandles = localUsableAxes(figHandle)
axesHandles = findall(figHandle, 'Type', 'axes');
if isempty(axesHandles)
    return;
end

keepMask = true(size(axesHandles));
for iAxis = 1:numel(axesHandles)
    try
        axisTag = string(get(axesHandles(iAxis), 'Tag'));
        if any(strcmpi(axisTag, ["legend", "colorbar", "Colorbar"]))
            keepMask(iAxis) = false;
        end
    catch
        keepMask(iAxis) = false;
    end
    try
        xLimits = get(axesHandles(iAxis), 'XLim');
        keepMask(iAxis) = keepMask(iAxis) && isnumeric(xLimits) && numel(xLimits) == 2 && all(isfinite(xLimits));
    catch
        keepMask(iAxis) = false;
    end
end
axesHandles = axesHandles(keepMask);
end

function localLinkCompatibleXAxis(axesHandles)
if numel(axesHandles) < 2
    return;
end

xLimits = NaN(numel(axesHandles), 2);
for iAxis = 1:numel(axesHandles)
    try
        xLimits(iAxis, :) = get(axesHandles(iAxis), 'XLim');
    catch
    end
end

validMask = all(isfinite(xLimits), 2) & abs(diff(xLimits, 1, 2)) > eps;
axesHandles = axesHandles(validMask);
xLimits = xLimits(validMask, :);
if numel(axesHandles) < 2
    return;
end

% Link only groups that already share a common x-domain. This avoids linking
% unrelated dashboard panels such as histograms, maps, and operating-point
% scatter plots that happen to live in the same figure.
tolerance = max(1e-9, 1e-4 * max(abs(xLimits(:))));
used = false(size(axesHandles));
for iAxis = 1:numel(axesHandles)
    if used(iAxis)
        continue;
    end
    sameDomain = abs(xLimits(:, 1) - xLimits(iAxis, 1)) <= tolerance & ...
        abs(xLimits(:, 2) - xLimits(iAxis, 2)) <= tolerance;
    groupAxes = axesHandles(sameDomain);
    if numel(groupAxes) >= 2
        try
            linkaxes(groupAxes, 'x');
        catch
        end
    end
    used = used | sameDomain;
end
end
