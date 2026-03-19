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
script3 = fullfile(scriptsDir, 'eBus_Release_Check.m');
if ~localIsFolderOnPath(scriptsDir)
    addpath(scriptsDir);
end

bannerWidth = 118;
topBottomSep = repmat('=', 1, bannerWidth);
midSep = repmat('-', 1, bannerWidth);

fprintf('%s\n', topBottomSep);
localPrintCentered(localMakeBoldText('Welcome to the eBT Tool'), bannerWidth);
localPrintCentered('Your companion for DIVe Simulation exploration, validation and visualization', bannerWidth);
fprintf('%s\n', midSep);
localPrintCentered('Click the tool below to continue...', bannerWidth);
fprintf('%s\n', topBottomSep);
fprintf('\n');

localPrintScriptLink( ...
    'DIVe_KPI_Plots_Check', ...
    script1, ...
    'matlab:DIVe_KPI_Plots_Check', ...
    'Check KPI''s, Plots and Generate reports for one or multiple simulations. Do Root Cause Analysis.');
localPrintScriptLink( ...
    'Batch_DIVe_Sim_Processing', ...
    script2, ...
    'matlab:Batch_DIVe_Sim_Processing', ...
    'Check KPI''s, Plots and Generate reports for a batch of simulations. Do Comparative Analysis.');
localPrintScriptLink( ...
    'eBus_Release_Check', ...
    script3, ...
    'matlab:eBus_Release_Check', ...
    'Compare eBus Configuration Releases. Both the Release folders must contain an identical names of set of .mat simulation files');
fprintf('\n');
localPrintThankYouBanner(bannerWidth);
end

function localPrintScriptLink(label, scriptPath, hrefCmd, descriptionText)
bullet = char(9670); % solid diamond
if ~isfile(scriptPath)
    fprintf('%s %s (missing: %s)\n', bullet, label, scriptPath);
    fprintf('\t%s\n\n', descriptionText);
    return;
end

fprintf('%s <a href="%s">%s</a>\n', bullet, hrefCmd, label);
fprintf('\t%s\n\n', descriptionText);
end

function localPrintCentered(textValue, totalWidth)
textValue = char(string(textValue));
padCount = max(0, floor((totalWidth - strlength(string(textValue))) / 2));
fprintf('%s%s\n', repmat(' ', 1, padCount), textValue);
end

function tf = localIsFolderOnPath(folderPath)
pathParts = strsplit(path, pathsep);
tf = any(strcmpi(pathParts, folderPath));
end

function localPrintThankYouBanner(bannerWidth)
sep = repmat('=', 1, bannerWidth);
message = 'Thank You';
fprintf('%s\n', sep);
localPrintCentered(localMakeBoldText(message), bannerWidth);
fprintf('%s\n', sep);
end

function out = localMakeBoldText(inText)
out = char(string(inText));
if ~localSupportsAnsiStyles()
    return;
end
esc = char(27);
out = sprintf('%s[1m%s%s[0m', esc, out, esc);
end

function tf = localSupportsAnsiStyles()
tf = false;
releaseTag = regexp(version('-release'), '\d{4}[ab]', 'match', 'once');
if isempty(releaseTag)
    return;
end
yearNum = str2double(releaseTag(1:4));
halfTag = lower(releaseTag(5));
tf = (yearNum > 2025) || (yearNum == 2025 && (halfTag == 'a' || halfTag == 'b'));
end
