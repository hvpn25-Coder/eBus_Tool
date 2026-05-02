function RCA_OpenReviewFigure(filePath, figureName)
% RCA_OpenReviewFigure  Open an RCA review plot from a Command Window link.
%
% The RCA report stores PNG files for Word/report embedding and now stores
% MATLAB FIG files beside them for interactive review. This function opens
% the FIG when available, otherwise it displays the PNG as an image fallback.

if nargin < 1 || strlength(string(filePath)) == 0
    warning('RCA_OpenReviewFigure:MissingFile', 'No figure file was provided.');
    return;
end
if nargin < 2 || strlength(string(figureName)) == 0
    figureName = "RCA Review Figure";
end

filePath = string(filePath);
figureName = char(string(figureName));

[folderPath, baseName, extension] = fileparts(char(filePath));
if isempty(folderPath)
    folderPath = pwd;
end

candidateFig = fullfile(folderPath, [baseName '.fig']);
if strcmpi(extension, '.fig')
    candidateFig = char(filePath);
end

if isfile(candidateFig)
    try
        fig = openfig(candidateFig, 'new', 'visible');
        set(fig, 'Name', figureName, 'NumberTitle', 'off', 'Color', 'w');
        RCA_PrepareInteractiveFigure(fig);
        drawnow;
        return;
    catch openException
        warning('RCA_OpenReviewFigure:FigOpenFailed', ...
            'Could not open MATLAB figure %s: %s. Falling back to image if available.', ...
            candidateFig, openException.message);
    end
end

if ~isfile(char(filePath))
    warning('RCA_OpenReviewFigure:FileNotFound', 'Figure file was not found: %s', char(filePath));
    return;
end

try
    imageData = imread(char(filePath));
    fig = figure('Color', 'w', 'Name', [figureName ' (image fallback)'], 'NumberTitle', 'off');
    ax = axes('Parent', fig);
    image(ax, imageData);
    axis(ax, 'image');
    axis(ax, 'off');
    title(ax, figureName, 'Interpreter', 'none');
    RCA_PrepareInteractiveFigure(fig);
    drawnow;
catch imageException
    warning('RCA_OpenReviewFigure:ImageOpenFailed', ...
        'Could not display figure %s: %s', char(filePath), imageException.message);
end
end
