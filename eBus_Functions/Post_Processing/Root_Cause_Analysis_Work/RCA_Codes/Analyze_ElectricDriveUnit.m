function result = Analyze_ElectricDriveUnit(analysisData, outputPaths, config)
% Analyze_ElectricDriveUnit  Integrated motor, gearbox, and final-drive RCA.
%
% The Electric Drive Unit (EDU) combines the electric machines/inverters,
% gearbox/transmission, and final-drive force path into one subsystem-level
% RCA owner view. The purpose is to evaluate the complete conversion chain:
% electrical power -> motor shaft torque -> gearbox output -> tractive force.

result = localInitResult("ELECTRIC DRIVE UNIT", ...
    {'emot1_act_trq', 'emot2_act_trq', 'emot1_act_spd', 'emot2_act_spd', 'emot1_pwr', 'emot2_pwr', 'gbx_out_trq', 'net_trac_trq'}, ...
    {'emot1_loss_pwr', 'emot2_loss_pwr', 'gbx_pwr_loss', 'gbx_out_spd', 'gr_num', 'gr_ratio', 'veh_long_force', 'whl_force'});

d = analysisData.Derived;
t = d.time_s(:);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Electric Drive Unit analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.SummaryText = summary;
    result.Suggestions = RCA_MakeSuggestionTable("Electric Drive Unit", recs, evidence);
    return;
end

unitFolder = fullfile(outputPaths.FiguresSubsystem, 'ElectricDriveUnit');
if ~exist(unitFolder, 'dir')
    mkdir(unitFolder);
end

% Run legacy module analyses internally so the combined EDU result preserves
% mature motor, transmission, and final-drive KPIs/plots without exposing
% three separate subsystem owners in the top-level RCA flow.
subResults = [ ...
    localRunLegacySubmodule(@Analyze_ElectricDrive, "Electric Drive", analysisData, outputPaths, config); ...
    localRunLegacySubmodule(@Analyze_Transmission, "Transmission", analysisData, outputPaths, config); ...
    localRunLegacySubmodule(@Analyze_FinalDrive, "Final Drive", analysisData, outputPaths, config)];

motorElecPower = d.motorElectricalPower_kW(:);
motorMechPower = d.motorMechanicalPower_kW(:);
motorLossPower = d.motorLossPower_kW(:);
gearboxLossPower = d.gearboxLossPower_kW(:);
gbxOutTrq = d.gearboxOutputTorque_Nm(:);
gbxOutSpd = d.gearboxOutputSpeed_rads(:);
tractionForce = d.tractionForce_N(:);
tractionPower = d.tractionPower_kW(:);
gear = d.gearNumber(:);
gearRatio = d.gearRatio(:);
motorSpeed = d.motorSpeed_rpm(:);
motorTorque = d.torqueActualTotal_Nm(:);
posTorqueLimit = d.torquePositiveLimit_Nm(:);
negTorqueLimit = d.torqueNegativeLimit_Nm(:);
eduEvidenceAvailable = any(isfinite(motorElecPower)) || any(isfinite(motorMechPower)) || ...
    any(isfinite(gbxOutTrq)) || any(isfinite(tractionForce));

gearboxOutputPower = gbxOutTrq .* gbxOutSpd / 1000;
if ~any(isfinite(gearboxOutputPower))
    gearboxOutputPower = tractionPower + gearboxLossPower;
end

elecDriveEnergy = RCA_TrapzFinite(t, max(motorElecPower, 0)) / 3600;
elecRegenEnergy = RCA_TrapzFinite(t, max(-motorElecPower, 0)) / 3600;
motorMechDriveEnergy = RCA_TrapzFinite(t, max(motorMechPower, 0)) / 3600;
gbxDriveEnergy = RCA_TrapzFinite(t, max(gearboxOutputPower, 0)) / 3600;
roadDriveEnergy = RCA_TrapzFinite(t, max(tractionPower, 0)) / 3600;
roadRegenEnergy = RCA_TrapzFinite(t, max(-tractionPower, 0)) / 3600;
motorLossEnergy = RCA_TrapzFinite(t, max(motorLossPower, 0)) / 3600;
gearboxLossEnergy = RCA_TrapzFinite(t, max(gearboxLossPower, 0)) / 3600;
unitLossEnergy = motorLossEnergy + gearboxLossEnergy;

driveEfficiency = 100 * roadDriveEnergy / max(elecDriveEnergy, eps);
shaftToRoadEfficiency = 100 * roadDriveEnergy / max(motorMechDriveEnergy, eps);
regenRecovery = 100 * elecRegenEnergy / max(roadRegenEnergy, eps);
unitLossShare = 100 * unitLossEnergy / max(elecDriveEnergy, eps);

driveMask = isfinite(motorTorque) & motorTorque > config.Thresholds.ControllerTorqueDeadband_Nm;
regenMask = isfinite(motorTorque) & motorTorque < -config.Thresholds.ControllerTorqueDeadband_Nm;
validPosLimit = driveMask & isfinite(posTorqueLimit) & posTorqueLimit > 0;
validNegLimit = regenMask & isfinite(negTorqueLimit) & abs(negTorqueLimit) > 0;
nearPositiveLimit = validPosLimit & motorTorque >= config.Thresholds.LimitUsageFraction .* posTorqueLimit;
nearRegenLimit = validNegLimit & abs(motorTorque) >= config.Thresholds.LimitUsageFraction .* abs(negTorqueLimit);

shiftMask = localShiftMask(gear);
shiftCount = sum(localShiftEvents(gear));
lossPowerTotal = max(motorLossPower, 0) + max(gearboxLossPower, 0);
driveLossShare = NaN(size(t));
drivePowerPositive = max(motorElecPower, 0);
driveLossShare(drivePowerPositive > eps) = 100 * lossPowerTotal(drivePowerPositive > eps) ./ ...
    max(drivePowerPositive(drivePowerPositive > eps), eps);

forceConsistencyResidual = NaN(size(t));
if any(isfinite(tractionForce)) && any(isfinite(gbxOutTrq)) && isfinite(d.finalDriveRatio) && isfinite(d.wheelRadius_m) && d.wheelRadius_m > eps
    estimatedForce = gbxOutTrq .* d.finalDriveRatio ./ d.wheelRadius_m;
    forceConsistencyResidual = tractionForce - estimatedForce;
end

rows = RCA_AddKPI(rows, 'EDU Electrical Drive Energy', elecDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'emot1_pwr + emot2_pwr', ...
    'Electrical drive-positive energy entering the electric machines and inverters.');
rows = RCA_AddKPI(rows, 'EDU Electrical Regen Energy', elecRegenEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'emot1_pwr + emot2_pwr', ...
    'Electrical recuperation energy returned from the EDU toward the HV system.');
rows = RCA_AddKPI(rows, 'EDU Motor Shaft Drive Energy', motorMechDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'motor torque + motor speed', ...
    'Positive mechanical energy at motor shafts before gearbox losses.');
rows = RCA_AddKPI(rows, 'EDU Gearbox Output Drive Energy', gbxDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'gbx_out_trq + gbx_out_spd', ...
    'Positive mechanical energy leaving the gearbox toward the final drive.');
rows = RCA_AddKPI(rows, 'EDU Road Tractive Drive Energy', roadDriveEnergy, 'kWh', ...
    'Energy Path', 'Electric Drive Unit', 'tractive force + vehicle speed', ...
    'Positive tractive energy delivered at the road interface.');
rows = RCA_AddKPI(rows, 'EDU Motor/Inverter Loss Energy', motorLossEnergy, 'kWh', ...
    'Losses', 'Electric Drive Unit', 'emot1_loss_pwr + emot2_loss_pwr', ...
    'Integrated motor and inverter loss energy.');
rows = RCA_AddKPI(rows, 'EDU Gearbox Loss Energy', gearboxLossEnergy, 'kWh', ...
    'Losses', 'Electric Drive Unit', 'gbx_pwr_loss', ...
    'Integrated gearbox/transmission loss energy.');
rows = RCA_AddKPI(rows, 'EDU Total Internal Loss Energy', unitLossEnergy, 'kWh', ...
    'Losses', 'Electric Drive Unit', 'motor loss + gearbox loss', ...
    'Motor/inverter plus gearbox loss energy inside the combined Electric Drive Unit.');
rows = RCA_AddKPI(rows, 'EDU Electrical-to-Road Drive Efficiency', driveEfficiency, '%', ...
    'Efficiency', 'Electric Drive Unit', 'motor electrical energy + road tractive energy', ...
    'Approximate drive efficiency from motor electrical power to road tractive power.');
rows = RCA_AddKPI(rows, 'EDU Shaft-to-Road Drive Efficiency', shaftToRoadEfficiency, '%', ...
    'Efficiency', 'Electric Drive Unit', 'motor shaft energy + road tractive energy', ...
    'Approximate mechanical transfer efficiency from motor shafts through gearbox/final-drive path.');
rows = RCA_AddKPI(rows, 'EDU Road-to-Electrical Regen Recovery', regenRecovery, '%', ...
    'Efficiency', 'Electric Drive Unit', 'road regen energy + motor electrical regen', ...
    'Approximate recuperation recovery from road negative work to electrical regen energy.');
rows = RCA_AddKPI(rows, 'EDU Loss Share of Electrical Drive Energy', unitLossShare, '%', ...
    'Losses', 'Electric Drive Unit', 'loss energy / electrical drive energy', ...
    'Combined motor/inverter and gearbox loss share normalized by drive electrical energy.');
rows = RCA_AddKPI(rows, 'EDU Near Positive Torque Limit Share', 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit), '%', ...
    'Capability', 'Electric Drive Unit', 'actual torque + max available torque', ...
    sprintf('Actual positive torque above %.0f%% of available torque indicates propulsion envelope usage.', 100 * config.Thresholds.LimitUsageFraction));
rows = RCA_AddKPI(rows, 'EDU Near Regen Torque Limit Share', 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit), '%', ...
    'Capability', 'Electric Drive Unit', 'actual torque + min available torque', ...
    sprintf('Actual negative torque above %.0f%% of available regen torque indicates recuperation envelope usage.', 100 * config.Thresholds.LimitUsageFraction));
rows = RCA_AddKPI(rows, 'EDU Shift Count', shiftCount, 'count', ...
    'Gear Operation', 'Electric Drive Unit', 'gr_num', ...
    'Actual gear changes inside the combined unit.');
rows = RCA_AddKPI(rows, 'EDU High Loss Share Time', 100 * RCA_FractionTrue(driveLossShare >= config.Thresholds.HighLossShare_pct, isfinite(driveLossShare)), '%', ...
    'Losses', 'Electric Drive Unit', 'motor loss + gearbox loss + motor electrical power', ...
    sprintf('Drive samples above %.1f%% internal loss share are flagged as high-loss EDU operation.', config.Thresholds.HighLossShare_pct));
rows = RCA_AddKPI(rows, 'EDU Force Path Residual MAE', mean(abs(forceConsistencyResidual), 'omitnan'), 'N', ...
    'Consistency', 'Electric Drive Unit', 'gbx_out_trq + final drive ratio + wheel radius + tractive force', ...
    'Checks whether gearbox torque, final-drive ratio, wheel radius, and tractive force are mutually consistent.');

summary(end + 1) = sprintf(['Electric Drive Unit energy path: %.2f kWh electrical drive energy enters the machines, ', ...
    '%.2f kWh reaches the road as positive tractive work, and %.2f kWh is recorded as motor plus gearbox loss.'], ...
    elecDriveEnergy, roadDriveEnergy, unitLossEnergy);
summary(end + 1) = sprintf(['Electric Drive Unit efficiency summary: electrical-to-road drive efficiency is %.1f%%, ', ...
    'shaft-to-road mechanical efficiency is %.1f%%, and road-to-electrical regen recovery is %.1f%%.'], ...
    driveEfficiency, shaftToRoadEfficiency, regenRecovery);
summary(end + 1) = sprintf(['Gear and capability context: %d shifts were observed; near positive torque limit share is %.1f%% ', ...
    'and near regen torque limit share is %.1f%%.'], ...
    shiftCount, 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit), 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit));
if isfinite(mean(abs(forceConsistencyResidual), 'omitnan'))
    summary(end + 1) = sprintf('Final-drive force-path consistency residual MAE is %.0f N, based on gearbox torque, final-drive ratio, wheel radius, and tractive force.', ...
        mean(abs(forceConsistencyResidual), 'omitnan'));
end

if unitLossShare > config.Thresholds.HighLossShare_pct
    recs(end + 1) = "Review the combined motor/inverter and gearbox loss path; EDU internal losses are high relative to electrical drive energy.";
    evidence(end + 1) = sprintf('EDU loss share of electrical drive energy is %.1f%%.', unitLossShare);
end
if 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit) > 20
    recs(end + 1) = "Separate upstream torque-command issues from EDU propulsion capability limits; the unit frequently operates near the positive torque envelope.";
    evidence(end + 1) = sprintf('Near positive torque limit share is %.1f%%.', 100 * RCA_FractionTrue(nearPositiveLimit, validPosLimit));
end
if 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit) > 20
    recs(end + 1) = "Review recuperation torque envelope and gear-dependent regen capability; the EDU frequently operates near the negative torque limit.";
    evidence(end + 1) = sprintf('Near regen torque limit share is %.1f%%.', 100 * RCA_FractionTrue(nearRegenLimit, validNegLimit));
end
if shiftCount > 0 && mean(lossPowerTotal(shiftMask), 'omitnan') > 1.25 * mean(lossPowerTotal(~shiftMask), 'omitnan')
    recs(end + 1) = "Review shift torque handover and ratio transition logic; EDU losses are materially higher during shift activity.";
    evidence(end + 1) = sprintf('Mean EDU loss power during shifts is %.2f kW versus %.2f kW away from shifts.', ...
        mean(lossPowerTotal(shiftMask), 'omitnan'), mean(lossPowerTotal(~shiftMask), 'omitnan'));
end
if isfinite(mean(abs(forceConsistencyResidual), 'omitnan')) && mean(abs(forceConsistencyResidual), 'omitnan') > config.Thresholds.ForceBalanceResidualWarn_N
    recs(end + 1) = "Check final-drive ratio, wheel-radius specification, or logged tractive-force convention; force-path residual is above the configured review threshold.";
    evidence(end + 1) = sprintf('Force-path residual MAE is %.0f N.', mean(abs(forceConsistencyResidual), 'omitnan'));
end

plotFiles = localAppendPlotFile(plotFiles, localPlotEnergyPath(unitFolder, elecDriveEnergy, motorMechDriveEnergy, gbxDriveEnergy, roadDriveEnergy, elecRegenEnergy, roadRegenEnergy, motorLossEnergy, gearboxLossEnergy, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotTorqueForcePath(unitFolder, t, motorTorque, gbxOutTrq, tractionForce, gear, gearRatio, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotLossAndShift(unitFolder, t, motorElecPower, motorLossPower, gearboxLossPower, driveLossShare, shiftMask, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotOperatingMap(unitFolder, motorSpeed, motorTorque, gear, driveLossShare, config));

mergedKpis = RCA_FinalizeKPITable(rows);
mergedKpis = [mergedKpis; localMergeSubmoduleKPIs(subResults)];
mergedSuggestions = RCA_MakeSuggestionTable("Electric Drive Unit", recs, evidence);
mergedSuggestions = [mergedSuggestions; localMergeSubmoduleSuggestions(subResults)];

plotFiles = [plotFiles; localMergeSubmoduleFigures(subResults)];
plotFiles = localExistingFiles(plotFiles);

for iSub = 1:numel(subResults)
    if numel(string(subResults(iSub).SummaryText)) > 0
        legacySummary = "Submodule evidence - " + string(subResults(iSub).Name) + ": " + string(subResults(iSub).SummaryText(:));
        summary = [summary(:); legacySummary(:)]; %#ok<AGROW>
    end
    if numel(string(subResults(iSub).Warnings)) > 0
        legacyWarnings = string(subResults(iSub).Warnings(:));
        result.Warnings = [result.Warnings(:); legacyWarnings(:)]; %#ok<AGROW>
    end
end

result.Available = eduEvidenceAvailable || any([subResults.Available]);
result.KPITable = mergedKpis;
result.FigureFiles = plotFiles;
result.SummaryText = unique(summary(summary ~= ""));
result.Suggestions = mergedSuggestions;
end

function kpiTable = localMergeSubmoduleKPIs(subResults)
kpiTable = RCA_FinalizeKPITable([]);
for iSub = 1:numel(subResults)
    if ~istable(subResults(iSub).KPITable) || height(subResults(iSub).KPITable) == 0
        continue;
    end
    tableValue = subResults(iSub).KPITable;
    if ~ismember('Subsystem', tableValue.Properties.VariableNames)
        tableValue.Subsystem = repmat("Electric Drive Unit", height(tableValue), 1);
    end
    tableValue.Subsystem(:) = "Electric Drive Unit";
    tableValue.Category = string(subResults(iSub).Name) + " / " + string(tableValue.Category);
    kpiTable = [kpiTable; tableValue]; %#ok<AGROW>
end
end

function suggestionTable = localMergeSubmoduleSuggestions(subResults)
suggestionTable = RCA_MakeSuggestionTable("Electric Drive Unit", strings(0, 1), strings(0, 1));
for iSub = 1:numel(subResults)
    if ~isfield(subResults(iSub), 'Suggestions') || ~istable(subResults(iSub).Suggestions) || height(subResults(iSub).Suggestions) == 0
        continue;
    end
    tableValue = subResults(iSub).Suggestions;
    tableValue.Subsystem(:) = "Electric Drive Unit";
    tableValue.Recommendation = string(subResults(iSub).Name) + " evidence: " + string(tableValue.Recommendation);
    suggestionTable = [suggestionTable; tableValue]; %#ok<AGROW>
end
end

function figureFiles = localMergeSubmoduleFigures(subResults)
figureFiles = strings(0, 1);
for iSub = 1:numel(subResults)
    if isfield(subResults(iSub), 'FigureFiles')
        figureFiles = [figureFiles; string(subResults(iSub).FigureFiles(:))]; %#ok<AGROW>
    end
end
end

function files = localExistingFiles(files)
files = string(files(:));
files = files(files ~= "");
keep = false(size(files));
for iFile = 1:numel(files)
    keep(iFile) = isfile(char(files(iFile)));
end
files = files(keep);
end

function shiftEvents = localShiftEvents(gear)
shiftEvents = false(size(gear));
if numel(gear) < 2
    return;
end
valid = isfinite(gear(2:end)) & isfinite(gear(1:end - 1));
shiftEvents(2:end) = valid & abs(diff(gear)) > 0.05;
end

function shiftMask = localShiftMask(gear)
shiftMask = localShiftEvents(gear);
if any(shiftMask)
    shiftMask = shiftMask | [shiftMask(2:end); false] | [false; shiftMask(1:end - 1)];
end
end

function plotFile = localPlotEnergyPath(outputFolder, elecDriveEnergy, motorMechDriveEnergy, gbxDriveEnergy, roadDriveEnergy, elecRegenEnergy, roadRegenEnergy, motorLossEnergy, gearboxLossEnergy, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(1, 2, 1);
bar(categorical({'Electrical in', 'Motor shaft', 'Gearbox out', 'Road tractive'}), ...
    [elecDriveEnergy, motorMechDriveEnergy, gbxDriveEnergy, roadDriveEnergy], 'FaceColor', config.Plot.Colors.Vehicle);
ylabel('Energy [kWh]');
title('Drive Energy Conversion Path');
grid on;

subplot(1, 2, 2);
bar(categorical({'Road regen', 'Electrical regen', 'Motor loss', 'Gearbox loss'}), ...
    [roadRegenEnergy, elecRegenEnergy, motorLossEnergy, gearboxLossEnergy], 'FaceColor', config.Plot.Colors.Motor);
ylabel('Energy [kWh]');
title('Regen and Loss Energy');
grid on;

sgtitle('Electric Drive Unit Energy Path');
plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_EnergyPath', config));
close(fig);
end

function plotFile = localPlotTorqueForcePath(outputFolder, t, motorTorque, gbxOutTrq, tractionForce, gear, gearRatio, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, motorTorque, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, gbxOutTrq, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
ylabel('Torque [Nm]');
title('Motor Torque to Gearbox Output Torque');
legend({'Total motor torque', 'Gearbox output torque', 'Zero'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, tractionForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
ylabel('Force [N]');
title('Final Drive Tractive Force Output');
grid on;

subplot(3, 1, 3);
yyaxis left;
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
ylabel('Gear [-]');
yyaxis right;
plot(t, gearRatio, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
ylabel('Gear ratio [-]');
xlabel('Time [s]');
title('Gear State and Ratio Context');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_TorqueForcePath', config));
close(fig);
end

function plotFile = localPlotLossAndShift(outputFolder, t, motorElecPower, motorLossPower, gearboxLossPower, driveLossShare, shiftMask, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, max(motorElecPower, 0), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(-motorElecPower, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
ylabel('Power [kW]');
title('Motor Electrical Drive and Regen Power');
legend({'Drive power', 'Regen power'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, max(motorLossPower, 0), 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(gearboxLossPower, 0), 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
legendText = {'Motor/inverter loss', 'Gearbox loss'};
if any(shiftMask)
    plot(t(shiftMask), max(gearboxLossPower(shiftMask), 0), 'o', 'Color', config.Plot.Colors.Vehicle, 'MarkerSize', 4);
    legendText{end + 1} = 'Shift samples';
end
ylabel('Loss [kW]');
title('Motor/Inverter and Gearbox Loss Power');
legend(legendText, 'Location', 'best');
grid on;

subplot(3, 1, 3);
plot(t, driveLossShare, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.HighLossShare_pct, '--', 'Color', config.Plot.Colors.Neutral);
ylabel('Loss share [%]');
xlabel('Time [s]');
title('EDU Internal Loss Share During Drive');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_LossAndShift', config));
close(fig);
end

function plotFile = localPlotOperatingMap(outputFolder, motorSpeed, motorTorque, gear, driveLossShare, config)
plotFile = "";
valid = isfinite(motorSpeed) & isfinite(motorTorque);
if ~any(valid)
    return;
end
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(1, 2, 1);
gearColor = gear;
gearColor(~isfinite(gearColor)) = 0;
scatter(abs(motorSpeed(valid)), motorTorque(valid), 12, gearColor(valid), 'filled');
xlabel('Motor speed [rpm]');
ylabel('Motor torque [Nm]');
title('Motor Operating Map Colored by Gear');
cb = colorbar;
cb.Label.String = 'Gear';
grid on;

subplot(1, 2, 2);
validLoss = valid & isfinite(driveLossShare);
if any(validLoss)
    scatter(abs(motorSpeed(validLoss)), motorTorque(validLoss), 12, driveLossShare(validLoss), 'filled');
end
xlabel('Motor speed [rpm]');
ylabel('Motor torque [Nm]');
title('Operating Map Colored by EDU Loss Share');
cb = colorbar;
cb.Label.String = 'Loss share [%]';
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'ElectricDriveUnit_OperatingMap', config));
close(fig);
end

function plotFiles = localAppendPlotFile(plotFiles, plotFile)
if nargin < 2 || strlength(string(plotFile)) == 0
    return;
end
plotFiles(end + 1, 1) = string(plotFile);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end

function subResult = localRunLegacySubmodule(functionHandle, displayName, analysisData, outputPaths, config)
try
    subResult = functionHandle(analysisData, outputPaths, config);
catch subException
    subResult = localInitResult(displayName, {}, {});
    subResult.Warnings = "Legacy " + displayName + " sub-analysis was skipped inside Electric Drive Unit: " + string(subException.message);
end
end
