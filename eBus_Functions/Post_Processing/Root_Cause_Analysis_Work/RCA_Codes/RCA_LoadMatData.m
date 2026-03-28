function rawData = RCA_LoadMatData(matFilePath, config)
% RCA_LoadMatData  Load MAT content, inspect variables, and build a flat index.

if nargin < 2 || isempty(config)
    config = RCA_Config();
end

workspace = load(matFilePath);
inventory = whos('-file', matFilePath);
sizeText = arrayfun(@(x) mat2str(x.size), inventory, 'UniformOutput', false);
inventoryTable = table(string({inventory.name}'), string(sizeText(:)), [inventory.bytes]', ...
    string({inventory.class}'), 'VariableNames', {'Variable', 'Size', 'Bytes', 'Class'});

flatIndex = localFlattenWorkspace(workspace, config.General.MaxRecursiveDepth);
aliasMap = localBuildAliasMap(flatIndex, config.SignalFallback.PreferredContainers);
[defaultTime, defaultTimeSource] = localResolveDefaultTime(workspace, flatIndex, config);

rawData = struct();
rawData.MatFilePath = string(matFilePath);
rawData.Workspace = workspace;
rawData.InventoryTable = inventoryTable;
rawData.FlatIndex = flatIndex;
rawData.AliasMap = aliasMap;
rawData.DefaultTime = defaultTime;
rawData.DefaultTimeSource = string(defaultTimeSource);
end

function flatIndex = localFlattenWorkspace(workspace, maxDepth)
entries = localEmptyEntry();
entries(1) = [];

topFields = fieldnames(workspace);
for iField = 1:numel(topFields)
    entries = localAppendEntries(entries, workspace.(topFields{iField}), topFields{iField}, '', 0, maxDepth);
end

if isempty(entries)
    flatIndex = localEmptyEntry();
    flatIndex(1) = [];
else
    flatIndex = entries;
end
end

function entries = localAppendEntries(entries, value, pathText, parentPath, depth, maxDepth)
entry = localEmptyEntry();
entry.Path = string(pathText);
entry.ParentPath = string(parentPath);
entry.LeafName = string(localLeafName(pathText));
entry.Class = string(class(value));
entry.Size = string(mat2str(size(value)));
entry.IsLeaf = ~(isstruct(value) || (isobject(value) && ~localHasTimeDataProperties(value)));
entry.IsNumericLeaf = entry.IsLeaf && (isnumeric(value) || islogical(value));
entry.VectorLength = localValueVectorLength(value);
entry.IsTimeCandidate = localLooksLikeTime(value, entry.LeafName);
entries(end + 1) = entry;

if depth >= maxDepth
    return;
end

if isstruct(value) && isscalar(value)
    fields = fieldnames(value);
    for iField = 1:numel(fields)
        childPath = sprintf('%s.%s', pathText, fields{iField});
        entries = localAppendEntries(entries, value.(fields{iField}), childPath, pathText, depth + 1, maxDepth);
    end
elseif localHasTimeDataProperties(value)
    try
        dataValue = value.Data;
        timeValue = value.Time;
        entries = localAppendEntries(entries, dataValue, sprintf('%s.Data', pathText), pathText, depth + 1, maxDepth);
        entries = localAppendEntries(entries, timeValue, sprintf('%s.Time', pathText), pathText, depth + 1, maxDepth);
    catch
    end
end
end

function aliasMap = localBuildAliasMap(flatIndex, aliasNames)
aliasMap = struct();
for iAlias = 1:numel(aliasNames)
    alias = char(aliasNames{iAlias});
    candidateIdx = find(strcmpi(string({flatIndex.LeafName}), alias), 1, 'first');
    if ~isempty(candidateIdx)
        aliasMap.(matlab.lang.makeValidName(alias)) = string(flatIndex(candidateIdx).Path);
    else
        aliasMap.(matlab.lang.makeValidName(alias)) = "";
    end
end
end

function [timeVector, sourcePath] = localResolveDefaultTime(workspace, flatIndex, config)
timeVector = [];
sourcePath = "";

if isempty(flatIndex)
    return;
end

candidateMask = [flatIndex.IsTimeCandidate];
if ~any(candidateMask)
    return;
end

candidates = flatIndex(candidateMask);
preferredNames = string(config.SignalFallback.PreferredTimeNames(:));

bestIdx = [];
bestScore = -inf;
for iCandidate = 1:numel(candidates)
    score = double(candidates(iCandidate).VectorLength);
    if any(strcmpi(candidates(iCandidate).LeafName, preferredNames))
        score = score + 1e6;
    end
    if score > bestScore
        bestScore = score;
        bestIdx = iCandidate;
    end
end

if isempty(bestIdx)
    return;
end

sourcePath = candidates(bestIdx).Path;
value = localEvaluateWorkspacePath(workspace, sourcePath);
if isnumeric(value) || islogical(value)
    timeVector = double(value(:));
end
end

function value = localEvaluateWorkspacePath(workspace, pathText)
value = [];
if strlength(string(pathText)) == 0
    return;
end

workspaceFields = fieldnames(workspace);
for iField = 1:numel(workspaceFields)
    eval(sprintf('%s = workspace.(workspaceFields{iField});', workspaceFields{iField}));
end
try
    value = eval(char(pathText));
catch
    value = [];
end
end

function out = localLeafName(pathText)
parts = split(string(pathText), '.');
out = char(parts(end));
end

function tf = localHasTimeDataProperties(value)
tf = false;
try
    tf = isobject(value) && isprop(value, 'Data') && isprop(value, 'Time');
catch
    tf = false;
end
end

function len = localValueVectorLength(value)
if isnumeric(value) || islogical(value)
    if isvector(value)
        len = numel(value);
    else
        len = max(size(value, 1), size(value, 2));
    end
elseif localHasTimeDataProperties(value)
    try
        len = numel(value.Time);
    catch
        len = 0;
    end
else
    len = 0;
end
end

function tf = localLooksLikeTime(value, leafName)
nameHint = lower(string(leafName));
tf = false;
if isnumeric(value) || islogical(value)
    data = double(value(:));
    tf = isvector(value) && numel(data) >= 3 && all(isfinite(data)) && all(diff(data) >= 0) && ...
        (contains(nameHint, "time") || strcmp(nameHint, "t") || strcmp(nameHint, "tout") || ...
        strcmp(nameHint, "sec") || strcmp(nameHint, "secs") || strcmp(nameHint, "seconds"));
elseif localHasTimeDataProperties(value)
    tf = true;
end
end

function entry = localEmptyEntry()
entry = struct('Path', "", 'ParentPath', "", 'LeafName', "", 'Class', "", ...
    'Size', "", 'IsLeaf', false, 'IsNumericLeaf', false, 'VectorLength', 0, 'IsTimeCandidate', false);
end
