function [signalStore, specStore, signalPresenceTable, specPresenceTable, extractionLog] = RCA_CheckSignalPresence(metadata, rawData, config)
% RCA_CheckSignalPresence  Evaluate workbook expressions and build stores.

if nargin < 3 || isempty(config)
    config = RCA_Config();
end

[signalStore, signalPresenceTable, signalLog] = localResolveCatalogEntries(metadata.SignalCatalog, rawData, 'signal', config);
[specStore, specPresenceTable, specLog] = localResolveCatalogEntries(metadata.SpecCatalog, rawData, 'specification', config);
extractionLog = [signalLog; specLog];
end

function [store, presenceTable, extractionLog] = localResolveCatalogEntries(catalog, rawData, entryType, config)
store = struct();
presenceRows = cell(0, 10);
logRows = cell(0, 6);

for iRow = 1:height(catalog)
    row = catalog(iRow, :);
    [entry, bestMatch, method, note, success] = localTryExtractEntry(row, rawData, entryType, config);

    fieldName = matlab.lang.makeValidName(char(row.VariableName));
    store.(fieldName) = entry;

    if strcmpi(entryType, 'signal')
        isRequired = logical(row.IsRequired);
        if success
            status = "Present";
        elseif isRequired
            status = "Missing";
        else
            status = "Optional Missing";
        end
    else
        status = string(entry.Status);
        if strlength(status) == 0
            status = "Missing";
        end
        isRequired = false;
    end

    presenceRows(end + 1, :) = {string(row.VariableName), string(row.Description), string(row.Unit), ...
        status, string(row.Subsystem), string(localRequirementText(isRequired)), ...
        string(method), string(entry.ExpressionUsed), string(bestMatch), string(note)}; %#ok<AGROW>
    logRows(end + 1, :) = {string(entryType), string(row.VariableName), string(method), string(entry.ExpressionUsed), success, string(note)}; %#ok<AGROW>
end

presenceTable = cell2table(presenceRows, 'VariableNames', {'SignalName', 'Description', 'Unit', ...
    'Status', 'Subsystem', 'Requirement', 'Method', 'ExpressionUsed', 'BestMatch', 'Note'});
extractionLog = cell2table(logRows, 'VariableNames', {'EntryType', 'VariableName', 'Method', 'ExpressionUsed', 'Success', 'Note'});
end

function [entry, bestMatch, method, note, success] = localTryExtractEntry(row, rawData, entryType, config)
bestMatch = "";
method = "Unavailable";
success = false;

entry = localEmptyStoreEntry(row, entryType);
evaluations = row.Evaluations{1};
if isempty(evaluations)
    evaluations = {};
end

for iEval = 1:numel(evaluations)
    expr = string(evaluations{iEval});
    if strlength(strtrim(expr)) == 0
        continue;
    end
    [value, time, evalNote, ok] = localEvaluateExpression(expr, rawData, entryType);
    if ok
        entry = localPopulateEntry(entry, value, time, expr, "WorkbookEvaluation", evalNote, false);
        entry = localApplyExplicitTimeSignalRule(entry, row);
        method = "WorkbookEvaluation";
        note = entry.Note;
        success = true;
        return;
    end
end

[fallbackExpr, bestMatch, fallbackNote] = localSearchFallbackExpression(row, rawData, entryType, config);
if strlength(fallbackExpr) > 0
    [value, time, evalNote, ok] = localEvaluateExpression(fallbackExpr, rawData, entryType);
    if ok
        entry = localPopulateEntry(entry, value, time, fallbackExpr, "FallbackMatch", fallbackNote + " | " + evalNote, true);
        entry = localApplyExplicitTimeSignalRule(entry, row);
        method = "FallbackMatch";
        note = entry.Note;
        success = true;
        return;
    end
end

entry.Status = "Missing";
entry.Note = "No workbook evaluation or fallback match produced a usable value.";
note = entry.Note;
end

function entry = localPopulateEntry(entry, value, time, expressionUsed, method, note, approximate)
entry.Available = true;
entry.Status = "Present";
entry.Data = value;
entry.Time = time;
entry.ExpressionUsed = string(expressionUsed);
entry.Note = string(note);
entry.Approximate = logical(approximate);
entry.Confidence = localConfidenceText(method, time, approximate);
end

function entry = localApplyExplicitTimeSignalRule(entry, row)
if ~isfield(entry, 'Available') || ~entry.Available
    return;
end
if ~localIsTimeSignalRow(row)
    return;
end
if isempty(entry.Data) || ~(isnumeric(entry.Data) || islogical(entry.Data))
    return;
end

candidateTime = double(entry.Data(:));
if numel(candidateTime) < 2 || ~localIsMonotonic(candidateTime)
    return;
end

entry.Time = candidateTime;
entry.Data = candidateTime;
entry.Note = strtrim(entry.Note + " Explicit workbook time signal used as reference candidate.");
entry.Confidence = "High";
end

function [value, time, note, ok] = localEvaluateExpression(expression, rawData, entryType)
value = [];
time = [];
note = "";
ok = false;

workspace = rawData.Workspace;
workspaceFields = fieldnames(workspace);
for iField = 1:numel(workspaceFields)
    eval(sprintf('%s = workspace.(workspaceFields{iField});', workspaceFields{iField}));
end

aliasNames = fieldnames(rawData.AliasMap);
for iAlias = 1:numel(aliasNames)
    aliasPath = rawData.AliasMap.(aliasNames{iAlias});
    if strlength(aliasPath) == 0 || any(strcmp(workspaceFields, aliasNames{iAlias}))
        continue;
    end
    try
        eval(sprintf('%s = eval(char(aliasPath));', aliasNames{iAlias}));
    catch
    end
end

try
    candidate = eval(char(expression));
catch
    return;
end

[value, time, note, ok] = localUnwrapCandidate(candidate, expression, rawData, entryType);
end

function [value, time, note, ok] = localUnwrapCandidate(candidate, expression, rawData, entryType)
value = [];
time = [];
note = "";
ok = false;

if isnumeric(candidate) || islogical(candidate)
    value = double(candidate);
    ok = ~isempty(value);
elseif isstruct(candidate)
    if isfield(candidate, 'Data') && isfield(candidate, 'Time')
        value = double(candidate.Data);
        time = double(candidate.Time(:));
        ok = ~isempty(value);
    elseif isfield(candidate, 'data') && isfield(candidate, 'time')
        value = double(candidate.data);
        time = double(candidate.time(:));
        ok = ~isempty(value);
    else
        return;
    end
elseif isobject(candidate)
    try
        if isprop(candidate, 'Data') && isprop(candidate, 'Time')
            value = double(candidate.Data);
            time = double(candidate.Time(:));
            ok = ~isempty(value);
        end
    catch
        ok = false;
    end
end

if ~ok
    return;
end

if strcmpi(entryType, 'signal')
    value = squeeze(value);
    if ismatrix(value) && min(size(value)) == 1
        value = value(:);
    end
    if isempty(time)
        time = localFindRelatedTime(expression, numel(value), rawData);
        if ~isempty(time)
            note = "Time recovered from related signal container.";
        elseif ~isempty(rawData.DefaultTime) && numel(rawData.DefaultTime) == numel(value)
            time = rawData.DefaultTime(:);
            note = "Time matched to dataset default time vector.";
        elseif isscalar(value)
            note = "Scalar signal extracted without time basis.";
        else
            note = "Signal extracted without explicit time; alignment will use sample index if needed.";
        end
    end
else
    note = "Specification extracted as logged/static value.";
end
end

function time = localFindRelatedTime(expression, targetLength, rawData)
time = [];
sourceTokens = localExtractSourceTokens(expression, rawData);
for iToken = 1:numel(sourceTokens)
    token = regexprep(char(sourceTokens{iToken}), '\([^\)]*\)', '');
    tokenParts = split(string(token), '.');
    for iPart = numel(tokenParts):-1:1
        base = strjoin(tokenParts(1:iPart), '.');
        candidateNames = {'Time', 'time', 'tout', 't'};
        for iName = 1:numel(candidateNames)
            candidateExpr = sprintf('%s.%s', base, candidateNames{iName});
            [candidateTime, ok] = localTryEvaluateRaw(candidateExpr, rawData);
            if ok && isnumeric(candidateTime) && numel(candidateTime) == targetLength && localIsMonotonic(candidateTime)
                time = double(candidateTime(:));
                return;
            end
        end
    end
end

if ~isempty(rawData.DefaultTime) && numel(rawData.DefaultTime) == targetLength
    time = rawData.DefaultTime(:);
    return;
end

flatIndex = rawData.FlatIndex;
for iEntry = 1:numel(flatIndex)
    if flatIndex(iEntry).IsTimeCandidate && flatIndex(iEntry).VectorLength == targetLength
        [candidateTime, ok] = localTryEvaluateRaw(flatIndex(iEntry).Path, rawData);
        if ok && isnumeric(candidateTime) && localIsMonotonic(candidateTime)
            time = double(candidateTime(:));
            return;
        end
    end
end
end

function [value, ok] = localTryEvaluateRaw(expression, rawData)
value = [];
ok = false;
workspace = rawData.Workspace;
workspaceFields = fieldnames(workspace);
for iField = 1:numel(workspaceFields)
    eval(sprintf('%s = workspace.(workspaceFields{iField});', workspaceFields{iField}));
end
aliasNames = fieldnames(rawData.AliasMap);
for iAlias = 1:numel(aliasNames)
    aliasPath = rawData.AliasMap.(aliasNames{iAlias});
    if strlength(aliasPath) == 0 || any(strcmp(workspaceFields, aliasNames{iAlias}))
        continue;
    end
    try
        eval(sprintf('%s = eval(char(aliasPath));', aliasNames{iAlias}));
    catch
    end
end
try
    value = eval(char(expression));
    ok = true;
catch
    ok = false;
end
end

function [fallbackExpr, bestPath, note] = localSearchFallbackExpression(row, rawData, entryType, config)
fallbackExpr = "";
bestPath = "";
note = "";

flatIndex = rawData.FlatIndex;
candidateMask = [flatIndex.IsNumericLeaf];
if strcmpi(entryType, 'signal')
    candidateMask = candidateMask & [flatIndex.VectorLength] > 0;
end
flatCandidates = flatIndex(candidateMask);
if isempty(flatCandidates)
    return;
end

variableTokens = localTokenizeText(row.VariableName);
descriptionTokens = localTokenizeText(row.Description);
subsystemTokens = localTokenizeText(row.Subsystem);
targetKey = localNormalizeText(row.VariableName);

scores = zeros(numel(flatCandidates), 1);
for iCandidate = 1:numel(flatCandidates)
    pathText = string(flatCandidates(iCandidate).Path);
    pathKey = localNormalizeText(pathText);
    score = 0;
    if contains(pathKey, targetKey)
        score = score + 4;
    end
    for iTok = 1:numel(variableTokens)
        if contains(pathKey, variableTokens{iTok})
            score = score + 2;
        end
    end
    for iTok = 1:numel(descriptionTokens)
        if contains(pathKey, descriptionTokens{iTok})
            score = score + 1;
        end
    end
    for iTok = 1:numel(subsystemTokens)
        if contains(pathKey, subsystemTokens{iTok})
            score = score + 0.5;
        end
    end
    if strcmpi(entryType, 'signal') && contains(pathText, "logout")
        score = score + 1;
    end
    if strcmpi(entryType, 'specification') && contains(pathText, "sMP")
        score = score + 1;
    end
    scores(iCandidate) = score;
end

[sortedScores, order] = sort(scores, 'descend');
if isempty(sortedScores) || sortedScores(1) < config.SignalFallback.MinimumFallbackScore
    return;
end

bestPath = string(flatCandidates(order(1)).Path);
fallbackExpr = bestPath;
note = "Fallback matched directly to logged signal path.";
if numel(sortedScores) > 1 && (sortedScores(1) - sortedScores(2)) <= config.SignalFallback.AmbiguityMargin
    note = note + " Ambiguity warning: more than one candidate scored similarly.";
end
end

function sourceTokens = localExtractSourceTokens(expression, rawData)
tokenMatches = regexp(char(expression), '[A-Za-z]\w*(?:\.[A-Za-z]\w*(?:\([^\)]*\))?)*', 'match');
tokenMatches = string(tokenMatches(:));
workspaceNames = string(fieldnames(rawData.Workspace));
aliasNames = string(fieldnames(rawData.AliasMap));

keepMask = false(size(tokenMatches));
for iToken = 1:numel(tokenMatches)
    rootToken = extractBefore(tokenMatches(iToken) + ".", ".");
    keepMask(iToken) = any(strcmp(rootToken, workspaceNames)) || any(strcmp(rootToken, aliasNames));
end
sourceTokens = cellstr(tokenMatches(keepMask));
end

function tokens = localTokenizeText(textValue)
clean = lower(regexprep(string(textValue), '[^a-zA-Z0-9_ ]', ' '));
parts = split(clean);
parts = parts(strlength(parts) >= 3);
tokens = cellstr(unique(parts, 'stable'));
end

function normalized = localNormalizeText(textValue)
normalized = lower(regexprep(string(textValue), '[^a-zA-Z0-9]', ''));
end

function tf = localIsMonotonic(value)
value = double(value(:));
tf = numel(value) >= 3 && all(isfinite(value)) && all(diff(value) >= 0);
end

function requirementText = localRequirementText(isRequired)
if isRequired
    requirementText = "Required";
else
    requirementText = "Optional";
end
end

function confidence = localConfidenceText(method, time, approximate)
if strcmpi(method, 'WorkbookEvaluation') && ~isempty(time) && ~approximate
    confidence = "High";
elseif strcmpi(method, 'FallbackMatch') || approximate
    confidence = "Medium";
else
    confidence = "Low";
end
end

function entry = localEmptyStoreEntry(row, entryType)
entry = struct();
entry.Available = false;
entry.Name = string(row.VariableName);
entry.Description = string(row.Description);
entry.Unit = string(row.Unit);
entry.Subsystem = string(row.Subsystem);
entry.PositiveMeaning = localOptionalRowString(row, 'PositiveMeaning');
entry.NegativeMeaning = localOptionalRowString(row, 'NegativeMeaning');
entry.SignConventionText = localOptionalRowString(row, 'SignConventionText');
entry.Status = "Missing";
entry.Note = "";
entry.Data = [];
entry.Time = [];
entry.AlignedData = [];
entry.AlignedTime = [];
entry.ExpressionUsed = "";
entry.EntryType = string(entryType);
entry.Approximate = false;
entry.Confidence = "Low";
end

function value = localOptionalRowString(row, fieldName)
value = "";
try
    if ismember(fieldName, row.Properties.VariableNames)
        value = string(row.(fieldName));
        if ~isempty(value)
            value = value(1);
        end
    end
catch
    value = "";
end
end

function tf = localIsTimeSignalRow(row)
tf = false;
try
    if ismember('IsTimeSignal', row.Properties.VariableNames)
        tf = logical(row.IsTimeSignal(1));
        if tf
            return;
        end
    end
catch
end

descriptionKey = localNormalizeText(row.Description);
variableKey = localNormalizeText(row.VariableName);
unitKey = localNormalizeText(row.Unit);

hasTimeText = strcmp(descriptionKey, "time") || strcmp(variableKey, "time") || ...
    strcmp(variableKey, "timesim") || strcmp(variableKey, "simtime") || ...
    contains(descriptionKey, "time") || contains(variableKey, "time");
hasSecondUnit = any(strcmp(unitKey, ["s", "sec", "secs", "second", "seconds"]));
tf = hasTimeText && (hasSecondUnit || strcmp(descriptionKey, "time"));
end
