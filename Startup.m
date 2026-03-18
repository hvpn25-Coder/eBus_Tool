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
    localTryRemovePath(repoRoot);
end
addpath(repoRoot);
rehash;
drawnow;

fprintf('Startup: Quick Access item "%s" is configured.\n', entryLabel);
end

function localEnableLegacyFavorite(repoRoot, entryLabel)
favoriteCode = localBuildFavoriteCode(repoRoot);
favoriteIconName = 'favorite_command_E';

try
    fc = com.mathworks.mlwidgets.favoritecommands.FavoriteCommands.getInstance();
    localRemoveLegacyFavorite(fc, entryLabel);

    newFavorite = com.mathworks.mlwidgets.favoritecommands.FavoriteCommandProperties();
    newFavorite.setLabel(entryLabel);
    % Keep default category for better compatibility with MATLAB R2023b/R2024x.
    newFavorite.setCode(favoriteCode);
    newFavorite.setIsOnQuickToolBar(true);
    localTryShowFavoriteLabelOnToolbar(newFavorite);
    localTrySetFavoriteIcon(newFavorite, favoriteIconName);

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

function favoriteCode = localBuildFavoriteCode(repoRoot)
escapedRepoRoot = strrep(repoRoot, '''', '''''');
favoriteCode = sprintf( ...
    ['repoRoot=''%s''; pathParts=regexp(path,pathsep,''split''); ' ...
     'if exist(repoRoot,''dir'') && ~any(strcmpi(pathParts,repoRoot)), addpath(repoRoot); end; ' ...
     'rehash; feval(str2func(''openEbusTool''));'], ...
    escapedRepoRoot);
end

function localRemoveLegacyFavorite(fc, expectedLabel)
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
        if ~labelMatches
            continue;
        end

        categoryName = localGetCategoryLabel(category);
        localTryRemoveFavorite(fc, expectedLabel, categoryName);
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

function localTryRemovePath(folderPath)
try
    rmpath(folderPath);
catch
    % Ignore path removal failures and continue with addpath.
end
end

function localTryShowFavoriteLabelOnToolbar(favoriteObject)
try
    favoriteObject.setIsShowingLabelOnToolBar(true);
catch
    % This property is not available in all MATLAB releases.
end
end

function localTrySetFavoriteIcon(favoriteObject, iconName)
try
    favoriteObject.setIconName(iconName);
catch
    % Icon APIs differ across MATLAB releases.
end
end

function categoryName = localGetCategoryLabel(category)
categoryName = 'Favorite Commands';
try
    categoryName = char(string(category.getLabel()));
catch
    % Fall back to the default category label.
end
end

function localTryRemoveFavorite(fc, favoriteLabel, categoryName)
try
    fc.removeCommand(char(favoriteLabel), char(categoryName));
catch
    try
        fc.removeCommand(char(favoriteLabel), 'Favorite Commands');
    catch
        % Ignore removal failures and let addCommand handle duplicates.
    end
end
end
