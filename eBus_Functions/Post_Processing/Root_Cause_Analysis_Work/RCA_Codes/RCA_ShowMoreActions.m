function RCA_ShowMoreActions(resultsInput, actionGroup)
% RCA_ShowMoreActions  Print follow-up RCA navigation links after a review.

if nargin < 1
    resultsInput = [];
end
if nargin < 2 || isempty(actionGroup)
    actionGroup = "ALL";
end

results = localResolveResults(resultsInput);
actionGroup = upper(string(actionGroup));

fprintf('\nInteractive Actions:\n');
if actionGroup == "ALL" || actionGroup == "RCA"
    localPrintRcaReviewLinks();
end
if actionGroup == "ALL" || actionGroup == "WORD"
    localPrintWordReportLinks();
end
if any(actionGroup == ["DETAIL", "DETAILS", "KPI"])
    localPrintDetailLinks(results);
end
end

function localPrintRcaReviewLinks()
fprintf('\nRCA Review:\n');
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowVehicleReview(RCA_Results);', ...
    'Open Complete Vehicle RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''DRIVER'');', ...
    'Open Driver RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''ENVIRONMENT'');', ...
    'Open Environment RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''POWER TRAIN CONTROLLER'');', ...
    'Open Powertrain Controller RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''ELECTRIC DRIVE UNIT'');', ...
    'Open Electric Drive Unit RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''PNEUMATIC BRAKE SYSTEM'');', ...
    'Open Pneumatic Brake RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''VEHICLE DYNAMICS'');', ...
    'Open Vehicle Dynamics RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''BATTERY'');', ...
    'Open Battery RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''BATTERY MANAGEMENT SYSTEM'');', ...
    'Open Battery Management System RCA review'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowSubsystemReview(RCA_Results, ''AUXILIARY LOAD'');', ...
    'Open Auxiliary RCA review'));
end

function localPrintDetailLinks(results)
fprintf('\nKPI, Plot, and Deep Analysis:\n');
fprintf('  %s\n', localMatlabHyperlink( ...
    'RCA_ShowVehicleReview(RCA_Results);', ...
    'Open complete vehicle KPI / segment / root-cause dashboard'));

if isfield(results, 'Paths') && isstruct(results.Paths)
    localPrintFolderLink(results.Paths, 'Tables', 'Open exported KPI and RCA tables folder');
    localPrintFolderLink(results.Paths, 'FiguresVehicle', 'Open vehicle plots folder');
    localPrintFolderLink(results.Paths, 'FiguresSubsystem', 'Open subsystem plots folder');
    localPrintFolderLink(results.Paths, 'Root', 'Open complete RCA output folder');
end
end

function localPrintWordReportLinks()
fprintf('\nWord Report Generation:\n');
fprintf('  %s\n', localMatlabHyperlink( ...
    'Generate_eBus_RCA_Word_Report(RCA_Results, RCA_Results.Paths.Root, struct(''Language'',''EN''));', ...
    'Generate Word report (English)'));
fprintf('  %s\n', localMatlabHyperlink( ...
    'Generate_eBus_RCA_Word_Report(RCA_Results, RCA_Results.Paths.Root, struct(''Language'',''DE''));', ...
    'Generate Word report (German)'));
end

function localPrintFolderLink(pathsStruct, fieldName, labelText)
try
    if isfield(pathsStruct, fieldName) && strlength(string(pathsStruct.(fieldName))) > 0 && ...
            isfolder(char(string(pathsStruct.(fieldName))))
        folderPath = char(string(pathsStruct.(fieldName)));
        fprintf('  %s\n', localMatlabHyperlink(sprintf('winopen(''%s'');', localEscapeQuotes(folderPath)), labelText));
    end
catch
end
end

function results = localResolveResults(resultsInput)
results = [];

if isempty(resultsInput)
    try
        if evalin('base', 'exist(''RCA_Results'',''var'')')
            results = evalin('base', 'RCA_Results');
        end
    catch
    end
elseif isstruct(resultsInput)
    results = resultsInput;
elseif ischar(resultsInput) || (isstring(resultsInput) && isscalar(resultsInput))
    loaded = load(char(string(resultsInput)));
    if isfield(loaded, 'results')
        results = loaded.results;
    elseif isfield(loaded, 'RCA_Results')
        results = loaded.RCA_Results;
    end
end

if isempty(results) || ~isstruct(results)
    error('RCA_ShowMoreActions:MissingResults', ...
        'Could not resolve RCA_Results. Run Vehicle_Detailed_Analysis first or pass RCA_Results explicitly.');
end
end

function escapedText = localEscapeQuotes(textValue)
escapedText = strrep(char(string(textValue)), '''', '''''');
end

function hyperlinkText = localMatlabHyperlink(commandText, labelText)
hyperlinkText = sprintf('<a href="matlab:%s">%s</a>', commandText, labelText);
end
