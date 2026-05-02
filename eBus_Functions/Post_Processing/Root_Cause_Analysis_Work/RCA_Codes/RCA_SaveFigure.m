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
matlabFigurePath = fullfile(outputFolder, [safeName '.fig']);

set(figHandle, 'Color', 'w');
set(findall(figHandle, '-property', 'FontSize'), 'FontSize', config.Plot.FontSize);
RCA_PrepareInteractiveFigure(figHandle);

try
    exportgraphics(figHandle, figurePath, 'Resolution', 150);
catch
    saveas(figHandle, figurePath);
end

try
    savefig(figHandle, matlabFigurePath);
catch
    % PNG export is the report-critical artifact. FIG export is optional and
    % used by RCA review hyperlinks for interactive MATLAB inspection.
end
end
