function result = Analyze_FinalDrive(analysisData, outputPaths, config)
% Analyze_FinalDrive  Final-drive torque-to-force delivery and road-force RCA.

result = localInitResult("FINAL DRIVE", {'net_trac_trq'}, {'veh_long_force', 'gbx_out_trq', 'whl_force', 'gr_num', 'gr_ratio'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Final drive analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Final Drive", recs, evidence);
    result.SummaryText = summary;
    return;
end

tractionForce = d.tractionForce_N(:);
wheelForce = d.wheelForce_N(:);
tractionPower = d.tractionPower_kW(:);
gbxOutTrq = d.gearboxOutputTorque_Nm(:);
vehSpeed = d.vehVel_kmh(:);
roadSlope = d.roadSlope_pct(:);
rollForce = localDerivedVector(d, 'rollingResistanceForce_N', n);
gradeForce = localDerivedVector(d, 'gradeForce_N', n);
aeroForce = localDerivedVector(d, 'aeroDragForce_N', n);
gear = d.gearNumber(:);
gearRatio = d.gearRatio(:);
finalDriveRatio = d.finalDriveRatio;
wheelRadius = d.wheelRadius_m;

if ~any(isfinite(tractionForce))
    tractionForce = localAlignedSignal(analysisData.Signals, 'veh_long_force', n);
end
if ~any(isfinite(wheelForce))
    wheelForce = localAlignedSignal(analysisData.Signals, 'whl_force', n);
end
if ~any(isfinite(gbxOutTrq))
    gbxOutTrq = localAlignedSignal(analysisData.Signals, 'gbx_out_trq', n);
end

if all(isnan(tractionForce)) && ~all(isnan(wheelForce))
    tractionForce = wheelForce;
end
if all(isnan(wheelForce))
    wheelForce = tractionForce;
end

if all(isnan(tractionForce))
    result.Warnings(end + 1) = "Final drive analysis skipped because net tractive force is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Final Drive", recs, evidence);
    result.SummaryText = summary;
    return;
end

movingMask = isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh;
driveMask = isfinite(tractionForce) & tractionForce > 0;
regenMask = isfinite(tractionForce) & tractionForce < 0;
launchMask = movingMask & vehSpeed <= max(10, config.Thresholds.CreepSpeed_kmh * 2);
uphillMask = movingMask & isfinite(roadSlope) & roadSlope >= config.Thresholds.UphillSlope_pct;
steepUphillMask = movingMask & isfinite(roadSlope) & roadSlope >= config.Thresholds.SteepSlope_pct;

forceMismatch = tractionForce - wheelForce;
tractionToWheelMismatch = abs(forceMismatch);

outputForceFromTorque = NaN(size(gbxOutTrq));
if isfinite(finalDriveRatio) && isfinite(wheelRadius) && wheelRadius > eps
    outputForceFromTorque = gbxOutTrq .* finalDriveRatio ./ wheelRadius;
end
forceTransferError = NaN(size(tractionForce));
validTransfer = isfinite(outputForceFromTorque) & isfinite(tractionForce) & abs(outputForceFromTorque) > 1;
forceTransferError(validTransfer) = 100 * abs(tractionForce(validTransfer) - outputForceFromTorque(validTransfer)) ./ max(abs(outputForceFromTorque(validTransfer)), eps);

totalRoadLoad = rollForce + gradeForce + aeroForce;
tractiveMargin = tractionForce - totalRoadLoad;
positiveMarginMask = movingMask & driveMask & isfinite(totalRoadLoad);

peakTractionForce = max(tractionForce, [], 'omitnan');
peakRegenForce = max(-tractionForce, [], 'omitnan');
launchForce = mean(tractionForce(launchMask & driveMask), 'omitnan');
uphillForce = mean(tractionForce(uphillMask & driveMask), 'omitnan');
steepUphillMargin = mean(tractiveMargin(steepUphillMask & driveMask), 'omitnan');

tractiveEnergy = RCA_TrapzFinite(t, max(tractionPower, 0)) / 3600;
regenEnergy = RCA_TrapzFinite(t, max(-tractionPower, 0)) / 3600;

rows = RCA_AddKPI(rows, 'Peak Tractive Force', peakTractionForce, 'N', ...
    'Performance', 'Final Drive', 'net_trac_trq', ...
    'Maximum positive delivered tractive force.');
rows = RCA_AddKPI(rows, 'Peak Recuperation Force', peakRegenForce, 'N', ...
    'Performance', 'Final Drive', 'net_trac_trq', ...
    'Maximum negative tractive force magnitude at the road.');
rows = RCA_AddKPI(rows, 'Average Positive Tractive Power', mean(max(tractionPower, 0), 'omitnan'), 'kW', ...
    'Performance', 'Final Drive', 'tractive force + vehicle speed', ...
    'Mean positive road power delivered through the final drive.');
rows = RCA_AddKPI(rows, 'Tractive Energy', tractiveEnergy, 'kWh', ...
    'Energy', 'Final Drive', 'tractive force + vehicle speed', ...
    'Integrated positive tractive power.');
rows = RCA_AddKPI(rows, 'Recuperation Energy at Road', regenEnergy, 'kWh', ...
    'Energy', 'Final Drive', 'tractive force + vehicle speed', ...
    'Integrated negative tractive power magnitude.');
rows = RCA_AddKPI(rows, 'Mean Launch Tractive Force', launchForce, 'N', ...
    'Operation', 'Final Drive', 'net_trac_trq + vehicle speed', ...
    'Average positive tractive force during low-speed launch behavior.');
rows = RCA_AddKPI(rows, 'Mean Uphill Tractive Force', uphillForce, 'N', ...
    'Operation', 'Final Drive', 'net_trac_trq + road slope', ...
    'Average positive tractive force during uphill operation.');
rows = RCA_AddKPI(rows, 'Mean Steep-Uphill Force Margin', steepUphillMargin, 'N', ...
    'Capability', 'Final Drive', 'tractive force - total road load', ...
    'Positive values indicate the final drive still has net force margin on steep uphill segments.');
rows = RCA_AddKPI(rows, 'Traction to Wheel Force Mismatch', mean(tractionToWheelMismatch, 'omitnan'), 'N', ...
    'Consistency', 'Final Drive', 'net_trac_trq + whl_force', ...
    'Difference between net tractive force and wheel force outputs.');
rows = RCA_AddKPI(rows, 'Final-Drive Force Transfer Error', mean(forceTransferError, 'omitnan'), '%', ...
    'Consistency', 'Final Drive', 'gbx_out_trq + finalDriveRatio + wheelRadius + net_trac_trq', ...
    'Difference between force inferred from gearbox torque and logged net tractive force.');
rows = RCA_AddKPI(rows, 'Positive Force Margin Share', 100 * RCA_FractionTrue(tractiveMargin > 0, positiveMarginMask), '%', ...
    'Capability', 'Final Drive', 'tractive force - road load', ...
    'Share of positive-drive moving samples where tractive force exceeds combined road load.');
rows = RCA_AddKPI(rows, 'Negative Margin Share in Uphill', 100 * RCA_FractionTrue(tractiveMargin < 0, uphillMask & driveMask & isfinite(totalRoadLoad)), '%', ...
    'Capability', 'Final Drive', 'tractive force - road load + road slope', ...
    'Share of uphill drive samples where delivered tractive force does not overcome modeled road load.');

[uniqueGears, gearForceMean, gearMarginMean, gearTimeShare] = localPerGearForceStats(gear, tractionForce, tractiveMargin);
for iGear = 1:numel(uniqueGears)
    gearLabel = sprintf('Gear %.0f', uniqueGears(iGear));
    rows = RCA_AddKPI(rows, [gearLabel ' Time Share'], gearTimeShare(iGear), '%', ...
        'Per Gear', 'Final Drive', 'gr_num', ...
        'Share of valid gear samples spent in this gear.');
    rows = RCA_AddKPI(rows, [gearLabel ' Mean Tractive Force'], gearForceMean(iGear), 'N', ...
        'Per Gear', 'Final Drive', 'gr_num + net_trac_trq', ...
        'Average tractive force while this gear is active.');
    rows = RCA_AddKPI(rows, [gearLabel ' Mean Force Margin'], gearMarginMean(iGear), 'N', ...
        'Per Gear', 'Final Drive', 'gr_num + tractive margin', ...
        'Average road-load margin while this gear is active.');
end

summary(end + 1) = sprintf(['Final drive peak tractive force is %.0f N with %.2f kWh of positive tractive energy delivered to the road. ', ...
    'Peak recuperation force magnitude is %.0f N.'], peakTractionForce, tractiveEnergy, peakRegenForce);
summary(end + 1) = sprintf(['Road-load delivery context: mean uphill tractive force is %.0f N and mean steep-uphill force margin is %.0f N. ', ...
    'This shows whether the final drive preserves enough wheel force in demanding route segments.'], uphillForce, steepUphillMargin);
summary(end + 1) = sprintf(['Consistency checks: mean traction-to-wheel mismatch is %.0f N and mean torque-to-force transfer error is %.1f%%. ', ...
    'Large values indicate force-scaling or driveline plumbing issues.'], ...
    mean(tractionToWheelMismatch, 'omitnan'), mean(forceTransferError, 'omitnan'));

if mean(tractionToWheelMismatch, 'omitnan') > 500
    recs(end + 1) = "Check net tractive-force construction, brake-force subtraction, and wheel-force plumbing because traction and wheel forces diverge materially.";
    evidence(end + 1) = sprintf('Mean traction-to-wheel mismatch is %.0f N.', mean(tractionToWheelMismatch, 'omitnan'));
end
if isfinite(mean(forceTransferError, 'omitnan')) && mean(forceTransferError, 'omitnan') > 20
    recs(end + 1) = "Review final-drive ratio, wheel radius, and sign convention assumptions because gearbox torque does not translate into the logged tractive force consistently.";
    evidence(end + 1) = sprintf('Mean torque-to-force transfer error is %.1f%%.', mean(forceTransferError, 'omitnan'));
end
if 100 * RCA_FractionTrue(tractiveMargin < 0, uphillMask & driveMask & isfinite(totalRoadLoad)) > 15
    recs(end + 1) = "Investigate why uphill delivered force often falls below combined road load. This can limit gradeability or make poor performance appear upstream of the final drive.";
    evidence(end + 1) = sprintf('Negative uphill force-margin share is %.1f%%.', ...
        100 * RCA_FractionTrue(tractiveMargin < 0, uphillMask & driveMask & isfinite(totalRoadLoad)));
end
if isfinite(launchForce) && launchForce < 0.4 * peakTractionForce
    recs(end + 1) = "Review low-speed torque multiplication and launch calibration; launch tractive force is modest relative to peak available road force.";
    evidence(end + 1) = sprintf('Mean launch tractive force is %.0f N versus peak %.0f N.', launchForce, peakTractionForce);
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'FinalDrive');
plotFiles = localAppendPlotFile(plotFiles, localPlotOverview(figureFolder, t, tractionForce, wheelForce, tractionPower, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotForceTransfer(figureFolder, gbxOutTrq, outputForceFromTorque, tractionForce, gearRatio, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotRoadLoadMargin(figureFolder, t, tractionForce, totalRoadLoad, tractiveMargin, roadSlope, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotGearContext(figureFolder, vehSpeed, gear, tractionForce, uniqueGears, gearForceMean, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Final Drive", recs, evidence);
result.FigureFiles = plotFiles;
end

function signalData = localAlignedSignal(signalStore, signalName, n)
signal = RCA_GetSignalData(signalStore, signalName);
signalData = NaN(n, 1);
if signal.Available
    if ~isempty(signal.AlignedData)
        signalData = localResizeVector(signal.AlignedData, n);
    elseif ~isempty(signal.Data)
        signalData = localResizeVector(signal.Data, n);
    end
end
end

function vector = localResizeVector(dataValue, n)
vector = NaN(n, 1);
dataValue = double(dataValue(:));
count = min(numel(dataValue), n);
if count > 0
    vector(1:count) = dataValue(1:count);
end
end

function vector = localDerivedVector(derivedStruct, fieldName, n)
vector = NaN(n, 1);
if isfield(derivedStruct, fieldName)
    vector = localResizeVector(derivedStruct.(fieldName), n);
end
end

function [uniqueGears, gearForceMean, gearMarginMean, gearTimeShare] = localPerGearForceStats(gear, tractionForce, tractiveMargin)
validGear = isfinite(gear);
uniqueGears = unique(round(gear(validGear)));
uniqueGears = uniqueGears(:)';
gearForceMean = NaN(size(uniqueGears));
gearMarginMean = NaN(size(uniqueGears));
gearTimeShare = NaN(size(uniqueGears));
for iGear = 1:numel(uniqueGears)
    gearMask = validGear & round(gear) == uniqueGears(iGear);
    gearForceMean(iGear) = mean(tractionForce(gearMask), 'omitnan');
    gearMarginMean(iGear) = mean(tractiveMargin(gearMask), 'omitnan');
    gearTimeShare(iGear) = 100 * RCA_FractionTrue(gearMask, validGear);
end
end

function plotFile = localPlotOverview(outputFolder, t, tractionForce, wheelForce, tractionPower, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, tractionForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, wheelForce, '--', 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Final Drive Force Delivery');
ylabel('Force (N)');
legend({'Net tractive force', 'Wheel force', 'Zero line'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, tractionPower, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Wheel-End Tractive Power');
ylabel('Power (kW)');
grid on;

subplot(3, 1, 3);
cumDrive = RCA_CumtrapzFinite(t, max(tractionPower, 0)) / 3600;
cumRegen = RCA_CumtrapzFinite(t, max(-tractionPower, 0)) / 3600;
plot(t, cumDrive, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, cumRegen, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Cumulative Road Energy');
xlabel('Time (s)');
ylabel('Energy (kWh)');
legend({'Positive tractive energy', 'Recuperation energy'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'FinalDrive_Overview', config));
close(fig);
end

function plotFile = localPlotForceTransfer(outputFolder, gbxOutTrq, outputForceFromTorque, tractionForce, gearRatio, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid = isfinite(gbxOutTrq) & isfinite(tractionForce);
scatter(gbxOutTrq(valid), tractionForce(valid), 12, config.Plot.Colors.Vehicle, 'filled');
title('Gearbox Output Torque Versus Net Tractive Force');
xlabel('Gearbox output torque (Nm)');
ylabel('Net tractive force (N)');
grid on;

subplot(2, 2, 2);
validTransfer = isfinite(outputForceFromTorque) & isfinite(tractionForce);
scatter(outputForceFromTorque(validTransfer), tractionForce(validTransfer), 12, config.Plot.Colors.Motor, 'filled'); hold on;
lims = axis;
minRef = min([lims(1), lims(3)]);
maxRef = max([lims(2), lims(4)]);
plot([minRef, maxRef], [minRef, maxRef], '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Inferred Versus Logged Tractive Force');
xlabel('Inferred force from gearbox torque (N)');
ylabel('Logged net tractive force (N)');
grid on;

subplot(2, 2, 3);
validRatio = isfinite(gearRatio) & isfinite(tractionForce);
scatter(gearRatio(validRatio), tractionForce(validRatio), 12, config.Plot.Colors.Gear, 'filled');
title('Tractive Force Versus Gear Ratio');
xlabel('Actual gear ratio (-)');
ylabel('Net tractive force (N)');
grid on;

subplot(2, 2, 4);
histogram(tractionForce(isfinite(tractionForce)), 40, 'FaceColor', config.Plot.Colors.Vehicle);
title('Tractive Force Distribution');
xlabel('Net tractive force (N)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'FinalDrive_ForceTransfer', config));
close(fig);
end

function plotFile = localPlotRoadLoadMargin(outputFolder, t, tractionForce, roadLoad, tractiveMargin, roadSlope, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, tractionForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, roadLoad, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Delivered Force Versus Combined Road Load');
ylabel('Force (N)');
legend({'Delivered tractive force', 'Combined road load'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, tractiveMargin, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Tractive Force Margin');
ylabel('Margin (N)');
grid on;

subplot(3, 1, 3);
plot(t, roadSlope, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
title('Road Slope');
xlabel('Time (s)');
ylabel('Slope (%)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'FinalDrive_RoadLoadMargin', config));
close(fig);
end

function plotFile = localPlotGearContext(outputFolder, vehSpeed, gear, tractionForce, uniqueGears, gearForceMean, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid = isfinite(vehSpeed) & isfinite(gear);
scatter(vehSpeed(valid), gear(valid), 12, tractionForce(valid), 'filled');
title('Gear Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Gear (-)');
cb1 = colorbar;
cb1.Label.String = 'Tractive force (N)';
grid on;

subplot(2, 2, 2);
valid2 = isfinite(vehSpeed) & isfinite(tractionForce);
scatter(vehSpeed(valid2), tractionForce(valid2), 12, gear(valid2), 'filled');
title('Tractive Force Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Tractive force (N)');
cb2 = colorbar;
cb2.Label.String = 'Gear (-)';
grid on;

subplot(2, 2, 3);
bar(uniqueGears, gearForceMean, 'FaceColor', config.Plot.Colors.Vehicle);
title('Mean Tractive Force by Gear');
xlabel('Gear (-)');
ylabel('Force (N)');
grid on;

subplot(2, 2, 4);
histogram(gear(isfinite(gear)), 'FaceColor', config.Plot.Colors.Gear);
title('Gear Usage Histogram');
xlabel('Gear (-)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'FinalDrive_GearContext', config));
close(fig);
end

function plotFiles = localAppendPlotFile(plotFiles, plotFile)
if strlength(plotFile) > 0
    plotFiles(end + 1, 1) = plotFile; %#ok<AGROW>
end
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
