function metadata = RCA_ReadSignalCatalog(excelFilePath, config)
% RCA_ReadSignalCatalog  Defensively parse workbook sheets and headers.

if nargin < 2 || isempty(config)
    config = RCA_Config();
end

try
    sheetList = sheetnames(excelFilePath);
catch
    [~, sheetList] = xlsfinfo(excelFilePath);
end

if isempty(sheetList)
    error('RCA_ReadSignalCatalog:NoSheets', 'No readable sheets were found in %s.', excelFilePath);
end

signalSheet = localSelectSheet(sheetList, ["signal", "evaluation"]);
specSheet = localSelectSheet(sheetList, ["spec"]);
blockSheet = localSelectSheet(sheetList, ["block", "diagram"]);

signalCatalog = localParseCatalogSheet(excelFilePath, signalSheet, 'signal', config);
specCatalog = localParseCatalogSheet(excelFilePath, specSheet, 'specification', config);
blockDiagram = localParseBlockSheet(excelFilePath, blockSheet);

if ~isempty(signalCatalog)
    signalCatalog.IsRequired = ismember(cellstr(signalCatalog.VariableName), config.RequiredSignalNames);
    signalCatalog.IsTimeSignal = localIdentifyTimeSignals(signalCatalog);
else
    signalCatalog.IsRequired = false(0, 1);
    signalCatalog.IsTimeSignal = false(0, 1);
end

subsystems = unique([signalCatalog.Subsystem; specCatalog.Subsystem; blockDiagram.Subsystem], 'stable');

metadata = struct();
metadata.ExcelFile = string(excelFilePath);
metadata.SheetNames = string(sheetList(:));
metadata.SignalSheet = string(signalSheet);
metadata.SpecSheet = string(specSheet);
metadata.BlockSheet = string(blockSheet);
metadata.SignalCatalog = signalCatalog;
metadata.SpecCatalog = specCatalog;
metadata.BlockDiagram = blockDiagram;
metadata.SubsystemList = subsystems;
metadata.TimeSignalNames = unique(signalCatalog.VariableName(signalCatalog.IsTimeSignal), 'stable');
end

function selectedSheet = localSelectSheet(sheetList, tokens)
sheetList = string(sheetList(:));
normalizedSheets = lower(regexprep(sheetList, '[^a-zA-Z0-9]', ''));
selectedSheet = sheetList(1);

for iSheet = 1:numel(sheetList)
    matched = true;
    for iToken = 1:numel(tokens)
        matched = matched && contains(normalizedSheets(iSheet), lower(regexprep(tokens(iToken), '[^a-zA-Z0-9]', '')));
    end
    if matched
        selectedSheet = sheetList(iSheet);
        return;
    end
end
end

function catalog = localParseCatalogSheet(excelFilePath, sheetName, mode, ~)
raw = readcell(excelFilePath, 'Sheet', char(sheetName));
headerRow = localFindHeaderRow(raw);
headers = string(raw(headerRow, :));
headerKeys = localNormalizeArray(headers);

subsystemCol = localFindColumn(headerKeys, ["subsystem"]);
descriptionCol = localFindColumn(headerKeys, ["description"]);
unitCol = localFindColumn(headerKeys, ["unit"]);
variableCol = localFindColumn(headerKeys, ["variable", "name"]);
evaluationCols = localFindEvaluationColumns(headerKeys, mode);

rows = cell(0, 9);
currentSubsystem = "";
for iRow = headerRow + 1:size(raw, 1)
    row = raw(iRow, :);
    if localRowIsEmpty(row)
        continue;
    end

    subsystem = localCellString(row{subsystemCol});
    if strlength(subsystem) > 0
        currentSubsystem = subsystem;
    end
    if strlength(currentSubsystem) == 0
        continue;
    end

    description = localCellString(row{descriptionCol});
    unit = localCellString(row{unitCol});
    variableName = localCellString(row{variableCol});
    evaluations = localCollectEvaluations(row, evaluationCols);

    if strlength(description) == 0 && strlength(variableName) == 0 && isempty(evaluations)
        continue;
    end
    if strlength(variableName) == 0
        variableName = matlab.lang.makeValidName(char(description));
    end

    [positiveMeaning, negativeMeaning, signConventionText] = localParseSignConvention(description);
    rows(end + 1, :) = {localPrettySubsystem(currentSubsystem), description, unit, variableName, evaluations, ...
        string(sheetName), positiveMeaning, negativeMeaning, signConventionText};
end

if isempty(rows)
    catalog = table(strings(0, 1), strings(0, 1), strings(0, 1), strings(0, 1), cell(0, 1), strings(0, 1), ...
        strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'Subsystem', 'Description', 'Unit', 'VariableName', 'Evaluations', 'SourceSheet', ...
        'PositiveMeaning', 'NegativeMeaning', 'SignConventionText'});
else
    catalog = cell2table(rows, 'VariableNames', {'Subsystem', 'Description', 'Unit', 'VariableName', 'Evaluations', 'SourceSheet', ...
        'PositiveMeaning', 'NegativeMeaning', 'SignConventionText'});
end
end

function blockDiagram = localParseBlockSheet(excelFilePath, sheetName)
raw = readcell(excelFilePath, 'Sheet', char(sheetName));
headerRow = localFindHeaderRow(raw);
headers = string(raw(headerRow, :));
headerKeys = localNormalizeArray(headers);

inputCol = localFindColumn(headerKeys, ["input"]);
subsystemCol = localFindColumn(headerKeys, ["subsystem"]);
outputCol = localFindColumn(headerKeys, ["output"]);

rows = cell(0, 4);
currentSubsystem = "";
for iRow = headerRow + 1:size(raw, 1)
    row = raw(iRow, :);
    if localRowIsEmpty(row)
        continue;
    end

    subsystem = localCellString(row{subsystemCol});
    if strlength(subsystem) > 0
        currentSubsystem = subsystem;
    end
    if strlength(currentSubsystem) == 0 || strcmpi(currentSubsystem, 'Subsystem')
        continue;
    end

    inputSignal = localCellString(row{inputCol});
    outputSignal = localCellString(row{outputCol});
    if strlength(inputSignal) == 0 && strlength(outputSignal) == 0
        continue;
    end

    rows(end + 1, :) = {localPrettySubsystem(currentSubsystem), inputSignal, outputSignal, string(sheetName)};
end

if isempty(rows)
    blockDiagram = table(strings(0, 1), strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'Subsystem', 'InputSignal', 'OutputSignal', 'SourceSheet'});
else
    blockDiagram = cell2table(rows, 'VariableNames', {'Subsystem', 'InputSignal', 'OutputSignal', 'SourceSheet'});
end
end

function headerRow = localFindHeaderRow(raw)
maxRows = min(size(raw, 1), 15);
bestScore = -inf;
headerRow = 1;
for iRow = 1:maxRows
    row = raw(iRow, :);
    textCells = cellfun(@(c) ischar(c) || isstring(c), row);
    nonEmpty = sum(~cellfun(@localIsEmptyCell, row));
    score = sum(textCells) + nonEmpty;
    if score > bestScore
        bestScore = score;
        headerRow = iRow;
    end
end
end

function idx = localFindColumn(headerKeys, tokens)
idx = 1;
for iCol = 1:numel(headerKeys)
    matched = true;
    for iToken = 1:numel(tokens)
        matched = matched && contains(headerKeys(iCol), localNormalizeScalar(tokens(iToken)));
    end
    if matched
        idx = iCol;
        return;
    end
end
end

function evaluationCols = localFindEvaluationColumns(headerKeys, mode)
evaluationCols = [];
for iCol = 1:numel(headerKeys)
    if contains(headerKeys(iCol), "evaluation")
        if strcmpi(mode, 'signal') && contains(headerKeys(iCol), "signal")
            evaluationCols(end + 1) = iCol; %#ok<AGROW>
        elseif strcmpi(mode, 'specification') && contains(headerKeys(iCol), "spec")
            evaluationCols(end + 1) = iCol; %#ok<AGROW>
        end
    end
end
if isempty(evaluationCols)
    evaluationCols = find(contains(headerKeys, "evaluation"));
end
end

function evaluations = localCollectEvaluations(row, evaluationCols)
evaluations = {};
for iCol = 1:numel(evaluationCols)
    value = localCellString(row{evaluationCols(iCol)});
    if strlength(value) > 0
        evaluations{end + 1, 1} = char(value); %#ok<AGROW>
    end
end
end

function flag = localRowIsEmpty(row)
flag = all(cellfun(@localIsEmptyCell, row));
end

function out = localPrettySubsystem(inputValue)
parts = split(lower(string(inputValue)));
parts(strlength(parts) == 0) = [];
parts = upper(extractBefore(parts, 2)) + extractAfter(parts, 1);
out = strjoin(parts, ' ');
end

function out = localCellString(value)
if isempty(value)
    out = "";
elseif isstring(value) || ischar(value)
    strValue = strtrim(string(value));
    if all(ismissing(strValue) | strlength(strValue) == 0)
        out = "";
    else
        out = strtrim(strValue(1));
    end
elseif isnumeric(value) || islogical(value)
    out = string(value);
else
    try
        strValue = strtrim(string(value));
        if all(ismissing(strValue) | strlength(strValue) == 0)
            out = "";
        else
            out = strtrim(strValue(1));
        end
    catch
        out = "";
    end
end
end

function tf = localIsEmptyCell(value)
if isempty(value)
    tf = true;
elseif isstring(value)
    tf = all(strlength(strtrim(value)) == 0 | ismissing(value));
elseif ischar(value)
    tf = isempty(strtrim(value));
elseif isnumeric(value) || islogical(value)
    tf = false;
else
    try
        missingMask = ismissing(value);
        tf = all(missingMask(:));
    catch
        tf = false;
    end
end
end

function normalized = localNormalizeArray(values)
normalized = strings(size(values));
for iValue = 1:numel(values)
    normalized(iValue) = localNormalizeScalar(values(iValue));
end
end

function normalized = localNormalizeScalar(value)
normalized = lower(regexprep(string(value), '[^a-zA-Z0-9]', ''));
end

function isTimeSignal = localIdentifyTimeSignals(signalCatalog)
isTimeSignal = false(height(signalCatalog), 1);
for iRow = 1:height(signalCatalog)
    descriptionKey = localNormalizeScalar(signalCatalog.Description(iRow));
    variableKey = localNormalizeScalar(signalCatalog.VariableName(iRow));
    unitKey = localNormalizeScalar(signalCatalog.Unit(iRow));

    hasTimeText = strcmp(descriptionKey, "time") || strcmp(variableKey, "time") || ...
        strcmp(variableKey, "timesim") || strcmp(variableKey, "simtime") || ...
        contains(descriptionKey, "time") || contains(variableKey, "time");
    hasSecondUnit = any(strcmp(unitKey, ["s", "sec", "secs", "second", "seconds"]));

    isTimeSignal(iRow) = hasTimeText && (hasSecondUnit || strcmp(descriptionKey, "time"));
end
end

function [positiveMeaning, negativeMeaning, signConventionText] = localParseSignConvention(description)
descriptionText = lower(string(description));
positiveMeaning = localExtractSignMeaning(descriptionText, 'positive');
negativeMeaning = localExtractSignMeaning(descriptionText, 'negative');

if strlength(positiveMeaning) == 0 && strlength(negativeMeaning) == 0
    signConventionText = "";
else
    signConventionText = "Positive=" + positiveMeaning + "; Negative=" + negativeMeaning;
end
end

function meaning = localExtractSignMeaning(descriptionText, polarityToken)
meaning = "";

switch lower(string(polarityToken))
    case "positive"
        if contains(descriptionText, "positive for")
            tokenText = extractAfter(descriptionText, "positive for");
        else
            tokenText = "";
        end
    case "negative"
        if contains(descriptionText, "negative for")
            tokenText = extractAfter(descriptionText, "negative for");
        else
            tokenText = "";
        end
    otherwise
        tokenText = "";
end

if strlength(tokenText) == 0
    return;
end

stopTokens = [" and ", ")", ",", ";"];
stopPos = strlength(tokenText) + 1;
for iToken = 1:numel(stopTokens)
    pos = strfind(char(tokenText), char(stopTokens(iToken)));
    if ~isempty(pos)
        stopPos = min(stopPos, pos(1));
    end
end
tokenText = strtrim(extractBefore(tokenText, stopPos));
tokenText = regexprep(tokenText, '\s+', ' ');

if contains(tokenText, "discharg")
    meaning = "discharge";
elseif contains(tokenText, "charg")
    meaning = "charge";
elseif contains(tokenText, "driv")
    meaning = "driving";
elseif contains(tokenText, "recuper") || contains(tokenText, "regen")
    meaning = "regeneration";
elseif contains(tokenText, "brak")
    meaning = "braking";
else
    meaning = strtrim(tokenText);
end
end
