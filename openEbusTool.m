function openEbusTool()
% Open eBus_Tool workspace folder and show startup links in MATLAB.
repoRoot = fileparts(mfilename('fullpath'));
if ~isfolder(repoRoot)
    warning('openEbusTool:MissingRepoRoot', 'Repository folder not found: %s', repoRoot);
    return;
end
cd(repoRoot);

scriptsDir = fullfile(repoRoot, 'eBus_Functions', 'Post_Processing', 'WorkScripts');
script1 = fullfile(scriptsDir, 'DIVe_KPI_Plots_Check.m');
script2 = fullfile(scriptsDir, 'Batch_DIVe_Sim_Processing.m');

bannerWidth = 118;
topBottomSep = repmat('=', 1, bannerWidth);
midSep = repmat('-', 1, bannerWidth);

fprintf('%s\n', topBottomSep);
localPrintCentered('Welcome to the eBT Tool', bannerWidth);
localPrintCentered('Your companion for DIVe Simulation exploration, validation and visualization', bannerWidth);
fprintf('%s\n', midSep);
localPrintCentered('Click the tool below to continue...', bannerWidth);
fprintf('%s\n', topBottomSep);
fprintf('\n');

localPrintScriptLink( ...
    'DIVe_KPI_Plots_Check', ...
    script1, ...
    'Check KPI''s, Plots and Generate reports for one or multiple simulations. Do Root Cause Analysis.');
localPrintScriptLink( ...
    'Batch_DIVe_Sim_Processing', ...
    script2, ...
    'Check KPI''s, Plots and Generate reports for a batch of simulations. Do Comparative Analysis.');
fprintf('\n');
end

function localPrintScriptLink(label, scriptPath, descriptionText)
bullet = char(9670); % solid diamond
if ~isfile(scriptPath)
    fprintf('%s %s (missing: %s)\n', bullet, label, scriptPath);
    fprintf('\t%s\n\n', descriptionText);
    return;
end

escapedPath = strrep(scriptPath, '''', '''''');
hrefCmd = sprintf('matlab:run(''%s'')', escapedPath);
fprintf('%s <a href="%s"><strong>%s</strong></a>\n', bullet, hrefCmd, label);
fprintf('\t%s\n\n', descriptionText);
end

function localPrintCentered(textValue, totalWidth)
textValue = char(string(textValue));
padCount = max(0, floor((totalWidth - strlength(string(textValue))) / 2));
fprintf('%s%s\n', repmat(' ', 1, padCount), textValue);
end
