function [workspaceOut, executionLog] = Run_Custom_PostProcessing_Codes(workspaceIn, customCodeDir, matFilePath)
% Run_Custom_PostProcessing_Codes  Execute user custom scripts on a MAT workspace.

workspaceOut = workspaceIn;
executionLog = strings(0, 1);

if nargin < 2 || strlength(string(customCodeDir)) == 0 || ~isfolder(char(string(customCodeDir)))
    return;
end
if nargin < 3
    matFilePath = "";
end

customCodeDir = char(string(customCodeDir));
scriptFiles = dir(fullfile(customCodeDir, '*.m'));
if isempty(scriptFiles)
    return;
end

for iFile = 1:numel(scriptFiles)
    scriptPath = fullfile(scriptFiles(iFile).folder, scriptFiles(iFile).name);
    if strcmpi(scriptFiles(iFile).name, 'Run_Custom_PostProcessing_Codes.m')
        continue;
    end
    if localIsFunctionFile(scriptPath)
        executionLog(end + 1, 1) = "Skipped function file: " + string(scriptFiles(iFile).name); %#ok<AGROW>
        continue;
    end

    try
        workspaceOut = localExecuteScript(workspaceOut, scriptPath);
        executionLog(end + 1, 1) = "Executed custom script: " + string(scriptFiles(iFile).name); %#ok<AGROW>
    catch ME
        executionLog(end + 1, 1) = "Custom script failed for " + string(matFilePath) + ...
            " [" + string(scriptFiles(iFile).name) + "]: " + string(ME.message); %#ok<AGROW>
    end
end
end

function tf = localIsFunctionFile(scriptPath)
tf = false;
try
    fileText = fileread(scriptPath);
catch
    return;
end

lines = regexp(fileText, '\r\n|\n|\r', 'split');
for iLine = 1:numel(lines)
    lineText = strtrim(string(lines{iLine}));
    if strlength(lineText) == 0
        continue;
    end
    if startsWith(lineText, "%") || startsWith(lineText, "%%")
        continue;
    end
    tf = startsWith(lineText, "function", 'IgnoreCase', true);
    return;
end
end

function workspaceOut = localExecuteScript(workspaceIn, scriptPath)
workspaceFields = fieldnames(workspaceIn);
for iField = 1:numel(workspaceFields)
    fieldName = workspaceFields{iField};
    if isvarname(fieldName)
        eval(sprintf('%s = workspaceIn.(workspaceFields{%d});', fieldName, iField));
    end
end

run(scriptPath);

excludeVars = ["workspaceIn", "workspaceOut", "workspaceFields", "iField", "fieldName", ...
    "scriptPath", "excludeVars"];
workspaceOut = localCollectWorkspace(excludeVars);
end

function workspaceOut = localCollectWorkspace(excludeVars)
workspaceOut = struct();
allVars = string(evalin('caller', 'who'));
keepVars = setdiff(allVars, excludeVars, 'stable');
for iVar = 1:numel(keepVars)
    varName = char(keepVars(iVar));
    workspaceOut.(varName) = evalin('caller', varName);
end
end
