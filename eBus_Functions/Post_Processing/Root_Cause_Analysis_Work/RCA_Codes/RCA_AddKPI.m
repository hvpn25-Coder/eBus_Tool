function rows = RCA_AddKPI(rows, kpiName, value, unit, category, subsystem, basis, note)
% RCA_AddKPI  Append one KPI row to a cell array.

if nargin < 1 || isempty(rows)
    rows = cell(0, 7);
end

if nargin < 8 || isempty(note)
    note = "Complete";
end
if nargin < 7 || isempty(basis)
    basis = "";
end
if nargin < 6 || isempty(subsystem)
    subsystem = "Vehicle";
end
if nargin < 5 || isempty(category)
    category = "General";
end
if nargin < 4 || isempty(unit)
    unit = "";
end

numericValue = NaN;
if isnumeric(value) || islogical(value)
    if isempty(value)
        numericValue = NaN;
        note = string(note) + " | Empty numeric value.";
    elseif isscalar(value)
        numericValue = double(value);
    else
        numericValue = double(value(1));
        note = string(note) + " | Non-scalar value reduced to first element for table storage.";
    end
else
    note = string(note) + " | Non-numeric value recorded in note: " + string(value);
end

rows(end + 1, :) = {string(kpiName), numericValue, string(unit), ...
    string(category), string(subsystem), string(basis), string(note)};
end
