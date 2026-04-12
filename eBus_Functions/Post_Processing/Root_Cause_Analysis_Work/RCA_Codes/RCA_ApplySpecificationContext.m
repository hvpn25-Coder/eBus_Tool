function result = RCA_ApplySpecificationContext(result, analysisData)
% RCA_ApplySpecificationContext  Append workbook specification context to subsystem RCA output.

if nargin < 2 || ~isstruct(result) || ~isstruct(analysisData)
    return;
end
if ~isfield(analysisData, 'Metadata') || ~isfield(analysisData, 'Specs') || ~isfield(result, 'Name')
    return;
end
if ~isfield(analysisData.Metadata, 'SpecCatalog') || isempty(analysisData.Metadata.SpecCatalog)
    return;
end

specCatalog = analysisData.Metadata.SpecCatalog;
subsystemMask = localMatchesSubsystem(specCatalog.Subsystem, string(result.Name));
if ~any(subsystemMask)
    return;
end

specSubset = specCatalog(subsystemMask, :);
specRows = cell(0, 7);
summaryLines = strings(0, 1);
missingLines = strings(0, 1);

for iRow = 1:height(specSubset)
    specRow = specSubset(iRow, :);
    specEntry = RCA_GetSignalData(analysisData.Specs, specRow.VariableName);

    if specEntry.Available && ~isempty(specEntry.Data)
        specData = specEntry.Data;
        if isnumeric(specData) || islogical(specData)
            if isscalar(specData)
                specRows = RCA_AddKPI(specRows, ...
                    "Configured " + string(specRow.Description), ...
                    double(specData(1)), string(specRow.Unit), ...
                    "Specification", localPrettyName(result.Name), ...
                    string(specRow.VariableName), ...
                    "Workbook specification loaded from the Specifications sheet.");
            else
                summaryLines(end + 1) = sprintf('Specification available: %s via %s [%s].', ...
                    char(string(specRow.Description)), char(string(specRow.VariableName)), mat2str(size(specData))); %#ok<AGROW>
            end
        else
            summaryLines(end + 1) = "Specification available: " + string(specRow.Description) + ...
                " via " + string(specRow.VariableName) + "."; %#ok<AGROW>
        end
    else
        missingLines(end + 1) = "Specification not resolved from MAT/context: " + ...
            string(specRow.Description) + " (" + string(specRow.VariableName) + ")."; %#ok<AGROW>
    end
end

specTable = RCA_FinalizeKPITable(specRows);
result.KPITable = localMergeKPITable(result.KPITable, specTable);

if ~isempty(summaryLines)
    prefix = "Specification context:";
    if ~ismember(prefix, string(result.SummaryText))
        result.SummaryText = [result.SummaryText(:); prefix; summaryLines(:)];
    else
        result.SummaryText = [result.SummaryText(:); summaryLines(:)];
    end
end

if ~isempty(missingLines)
    result.Warnings = [result.Warnings(:); missingLines(:)];
end
end

function tf = localMatchesSubsystem(subsystemValues, targetName)
subsystemValues = string(subsystemValues(:));
targetKey = localNormalizeName(targetName);
targetAliases = localSubsystemSpecAliases(targetKey);
tf = false(size(subsystemValues));
for iValue = 1:numel(subsystemValues)
    tf(iValue) = any(localNormalizeName(subsystemValues(iValue)) == targetAliases);
end
end

function aliases = localSubsystemSpecAliases(targetKey)
aliases = string(targetKey);
switch string(targetKey)
    case "ELECTRICDRIVEUNIT"
        aliases = [aliases, "ELECTRICDRIVE", "EDRIVE", "MOTORDRIVE", ...
            "TRANSMISSION", "GEARBOX", "FINALDRIVE", "DIFFERENTIAL", "AXLEDRIVE"];
end
aliases = unique(aliases);
end

function key = localNormalizeName(nameValue)
key = upper(regexprep(string(nameValue), '[^A-Za-z0-9]', ''));
end

function pretty = localPrettyName(nameValue)
pretty = regexprep(char(string(nameValue)), '([a-z])([A-Z])', '$1 $2');
pretty = strrep(pretty, '_', ' ');
pretty = strtrim(pretty);
if isempty(pretty)
    pretty = 'Subsystem';
end
end

function mergedTable = localMergeKPITable(existingTable, specTable)
if ~istable(existingTable) || isempty(existingTable)
    mergedTable = specTable;
    return;
end
if ~istable(specTable) || isempty(specTable)
    mergedTable = existingTable;
    return;
end

existingNames = string(existingTable.Properties.VariableNames);
specNames = string(specTable.Properties.VariableNames);
if ~isequal(existingNames, specNames)
    if ~ismember('Subsystem', existingNames) && ismember('Subsystem', specNames)
        specTable.Subsystem = [];
    elseif ismember('Subsystem', existingNames) && ~ismember('Subsystem', specNames)
        specTable = addvars(specTable, repmat("Subsystem", height(specTable), 1), 'Before', 5, 'NewVariableNames', 'Subsystem');
    end
end

specTable = specTable(:, existingTable.Properties.VariableNames);
mergedTable = [existingTable; specTable];
end
