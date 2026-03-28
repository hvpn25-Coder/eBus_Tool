function figurePath = RCA_SaveFigure(figHandle, outputFolder, baseName, config)
% RCA_SaveFigure  Save a figure as PNG using a safe file name.

if nargin < 4 || isempty(config)
    config = RCA_Config();
end

if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

safeName = regexprep(char(string(baseName)), '[^a-zA-Z0-9_\-]', '_');
figurePath = fullfile(outputFolder, [safeName '.png']);

set(figHandle, 'Color', 'w');
set(findall(figHandle, '-property', 'FontSize'), 'FontSize', config.Plot.FontSize);
axesHandles = findall(figHandle, 'Type', 'axes');
for iAxis = 1:numel(axesHandles)
    try
        axesHandles(iAxis).Toolbar = [];
    catch
    end
end

try
    exportgraphics(figHandle, figurePath, 'Resolution', 150);
catch
    saveas(figHandle, figurePath);
end
end
