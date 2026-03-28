function kpiTable = RCA_FinalizeKPITable(rows)
% RCA_FinalizeKPITable  Convert KPI cell rows into a standard table.

if nargin < 1 || isempty(rows)
    kpiTable = table(strings(0, 1), zeros(0, 1), strings(0, 1), strings(0, 1), ...
        strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'KPIName', 'Value', 'Unit', 'Category', ...
        'Subsystem', 'SignalBasis', 'StatusNote'});
    return;
end

kpiTable = cell2table(rows, 'VariableNames', {'KPIName', 'Value', 'Unit', ...
    'Category', 'Subsystem', 'SignalBasis', 'StatusNote'});
end
