% Startup.m
% Configures "eBus_Tool" in the MATLAB quick access area.

repoRoot = fileparts(mfilename('fullpath'));
entryLabel = 'eBus_Tool';

if ~usejava('desktop')
    fprintf('Startup: MATLAB desktop is not available. Skipping setup.\n');
    return;
end

if localIsR2025aOrNewer()
    localEnableQuickAccessExtension(repoRoot, entryLabel);
else
    localEnableLegacyFavorite(repoRoot, entryLabel);
end

function localEnableQuickAccessExtension(repoRoot, entryLabel)
requiredFiles = {
    fullfile(repoRoot, 'resources', 'extensions.json')
    fullfile(repoRoot, 'resources', 'icons', 'eBus_E.svg')
    fullfile(repoRoot, 'openEbusTool.m')
};

missing = requiredFiles(~cellfun(@isfile, requiredFiles));
if ~isempty(missing)
    warning('Startup:MissingQuickAccessFiles', ...
        'Quick access setup files are missing:\n%s', strjoin(string(missing), newline));
    return;
end

if localIsFolderOnPath(repoRoot)
    try
        rmpath(repoRoot);
    catch
    end
end
addpath(repoRoot);
rehash;
drawnow;

fprintf('Startup: Quick Access item "%s" is configured.\n', entryLabel);
end

function localEnableLegacyFavorite(repoRoot, entryLabel)
favoriteCode = 'openEbusTool;';
favoriteIconName = 'favorite_command_E';

try
    fc = com.mathworks.mlwidgets.favoritecommands.FavoriteCommands.getInstance();

    if localLegacyFavoriteExists(fc, entryLabel, favoriteCode)
        fprintf('Startup: Favorite "%s" already configured.\n', entryLabel);
        return;
    end

    newFavorite = com.mathworks.mlwidgets.favoritecommands.FavoriteCommandProperties();
    newFavorite.setLabel(entryLabel);
    % Keep default category for better compatibility with MATLAB R2023b/R2024x.
    newFavorite.setCode(favoriteCode);
    newFavorite.setIsOnQuickToolBar(true);

    try
        newFavorite.setIsShowingLabelOnToolBar(true);
    catch
    end
    try
        newFavorite.setIconName(favoriteIconName);
    catch
    end

    if ~localIsFolderOnPath(repoRoot)
        addpath(repoRoot);
    end
    fc.addCommand(newFavorite);
    drawnow;
    fprintf('Startup: Added Favorite "%s" to Quick Access Toolbar.\n', entryLabel);
catch ME
    warning('Startup:LegacyFavoriteSetupFailed', ...
        'Could not configure Favorite "%s": %s', entryLabel, ME.message);
end
end

function tf = localIsR2025aOrNewer()
tf = false;
releaseTag = regexp(version('-release'), '(\d{4})([ab])', 'tokens', 'once');
if isempty(releaseTag)
    return;
end

yearNum = str2double(releaseTag{1});
halfTag = lower(string(releaseTag{2}));
tf = (yearNum > 2025) || (yearNum == 2025 && (halfTag == "a" || halfTag == "b"));
end

function tf = localIsFolderOnPath(folderPath)
parts = strsplit(path, pathsep);
tf = any(strcmpi(parts, folderPath));
end

function tf = localLegacyFavoriteExists(fc, expectedLabel, expectedCode)
tf = false;
categories = localGetFavoriteCategories(fc);

if isempty(categories)
    return;
end

for iCat = 0:categories.size()-1
    category = categories.get(iCat);
    try
        children = category.getChildren();
    catch
        continue;
    end

    for iChild = 0:children.size()-1
        child = children.get(iChild);
        labelMatches = strcmp(string(child.getLabel()), string(expectedLabel));
        codeMatches = strcmp(strtrim(string(child.getCode())), strtrim(string(expectedCode)));
        if ~(labelMatches && codeMatches)
            continue;
        end

        try
            child.setIsOnQuickToolBar(true);
        catch
        end
        try
            child.setIconName('favorite_command_E');
        catch
        end
        tf = true;
        return;
    end
end
end

function categories = localGetFavoriteCategories(fc)
categories = [];
try
    method = fc.getClass().getDeclaredMethod('getCategories', javaArray('java.lang.Class', 0));
    method.setAccessible(true);
    categories = method.invoke(fc, javaArray('java.lang.Object', 0));
catch
    categories = [];
end
end
