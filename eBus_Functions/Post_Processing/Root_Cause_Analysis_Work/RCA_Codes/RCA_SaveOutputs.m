function RCA_SaveOutputs(results, outputPaths, config)
% RCA_SaveOutputs  Write tables, summary text, and MAT results to disk.

tablesToWrite = { ...
    'SignalPresence', results.SignalPresence; ...
    'SpecificationPresence', results.SpecPresence; ...
    'Thresholds', config.ThresholdTable; ...
    'MatInventory', results.MatInventory; ...
    'VehicleKPI', results.VehicleKPI; ...
    'SegmentKPI', results.SegmentKPI; ...
    'SegmentSummary', results.SegmentSummary; ...
    'RootCauseRanking', results.RootCauseRanking; ...
    'BadSegments', results.BadSegmentTable; ...
    'OptimizationSuggestions', results.OptimizationTable; ...
    'ExtractionLog', results.ExtractionLog};

for iTable = 1:size(tablesToWrite, 1)
    localSafeWriteTable(tablesToWrite{iTable, 2}, fullfile(outputPaths.Tables, [tablesToWrite{iTable, 1} '.csv']));
end

for iSubsystem = 1:numel(results.SubsystemResults)
    subsystemName = regexprep(char(results.SubsystemResults(iSubsystem).Name), '[^a-zA-Z0-9_\-]', '_');
    localSafeWriteTable(results.SubsystemResults(iSubsystem).KPITable, fullfile(outputPaths.Tables, ['Subsystem_' subsystemName '_KPI.csv']));
    localSafeWriteTable(results.SubsystemResults(iSubsystem).Suggestions, fullfile(outputPaths.Tables, ['Subsystem_' subsystemName '_Suggestions.csv']));
end

summaryFile = fullfile(outputPaths.Logs, 'RCA_Summary.txt');
fid = fopen(summaryFile, 'w');
if fid ~= -1
    fileCleanup = onCleanup(@() fclose(fid));
    fprintf(fid, 'Electric Bus Root Cause Analysis Summary\n');
    fprintf(fid, 'Generated on: %s\n\n', char(string(datetime('now'))));
    fprintf(fid, 'Vehicle Narrative\n');
    fprintf(fid, '-----------------\n');
    for iLine = 1:numel(results.VehicleNarrative)
        fprintf(fid, '%s\n', results.VehicleNarrative(iLine));
    end

    fprintf(fid, '\nRoot Cause Narrative\n');
    fprintf(fid, '--------------------\n');
    for iLine = 1:numel(results.RootCauseNarrative)
        fprintf(fid, '%s\n', results.RootCauseNarrative(iLine));
    end

    fprintf(fid, '\nVehicle Plot Notes\n');
    fprintf(fid, '------------------\n');
    for iLine = 1:numel(results.VehiclePlots.Notes)
        fprintf(fid, '%s\n', results.VehiclePlots.Notes(iLine));
    end

    fprintf(fid, '\nSubsystem Summaries\n');
    fprintf(fid, '-------------------\n');
    for iSubsystem = 1:numel(results.SubsystemResults)
        fprintf(fid, '\n[%s]\n', results.SubsystemResults(iSubsystem).Name);
        for iLine = 1:numel(results.SubsystemResults(iSubsystem).SummaryText)
            fprintf(fid, '%s\n', results.SubsystemResults(iSubsystem).SummaryText(iLine));
        end
        for iWarn = 1:numel(results.SubsystemResults(iSubsystem).Warnings)
            fprintf(fid, 'Warning: %s\n', results.SubsystemResults(iSubsystem).Warnings(iWarn));
        end
    end
    clear fileCleanup
end

save(fullfile(outputPaths.Root, 'RCA_Results.mat'), 'results', '-v7.3');
end

function localSafeWriteTable(tableValue, filePath)
try
    writetable(tableValue, filePath);
catch
    try
        writecell(table2cell(tableValue), filePath);
    catch
    end
end
end
