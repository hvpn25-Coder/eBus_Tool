function result = Analyze_Transmission(analysisData, outputPaths, config)
% Analyze_Transmission  Gearbox torque transfer, shift quality, and loss RCA.

result = localInitResult("TRANSMISSION", ...
    {'gbx_out_trq', 'gbx_out_spd', 'gr_num'}, ...
    {'gr_ratio', 'gbx_pwr_loss', 'emot1_act_trq', 'emot2_act_trq'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Transmission analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Transmission", recs, evidence);
    result.SummaryText = summary;
    return;
end

gear = d.gearNumber(:);
gearRatio = d.gearRatio(:);
gbxOutTrq = d.gearboxOutputTorque_Nm(:);
gbxOutSpd = d.gearboxOutputSpeed_rads(:);
gbxLoss = d.gearboxLossPower_kW(:);
vehSpeed = d.vehVel_kmh(:);
motorTrqIn = d.torqueActualTotal_Nm(:);
motorMechPwrIn = d.motorMechanicalPower_kW(:);

if ~any(isfinite(gear))
    gear = localAlignedSignal(analysisData.Signals, 'gr_num', n);
end
if ~any(isfinite(gearRatio))
    gearRatio = localAlignedSignal(analysisData.Signals, 'gr_ratio', n);
end
if ~any(isfinite(gbxOutTrq))
    gbxOutTrq = localAlignedSignal(analysisData.Signals, 'gbx_out_trq', n);
end
if ~any(isfinite(gbxOutSpd))
    gbxOutSpd = localAlignedSignal(analysisData.Signals, 'gbx_out_spd', n);
end
if ~any(isfinite(gbxLoss))
    gbxLoss = localAlignedSignal(analysisData.Signals, 'gbx_pwr_loss', n);
end

if ~any(isfinite(gear)) && ~any(isfinite(gbxOutTrq)) && ~any(isfinite(gbxOutSpd))
    result.Warnings(end + 1) = "Transmission analysis skipped because gear and gearbox output signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Transmission", recs, evidence);
    result.SummaryText = summary;
    return;
end

gbxOutPwr = gbxOutTrq .* gbxOutSpd / 1000;
driveMask = isfinite(motorMechPwrIn) & isfinite(gbxOutPwr) & motorMechPwrIn > 1 & gbxOutPwr > 0;
regenMask = isfinite(motorMechPwrIn) & isfinite(gbxOutPwr) & motorMechPwrIn < -1 & gbxOutPwr < 0;
movingMask = isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh;

shiftEvents = localShiftEvents(gear);
shiftMask = localShiftMask(gear);
shiftIdx = find(shiftEvents);
shiftCount = numel(shiftIdx);
distanceKm = max(d.tripDistance_km, eps);
shiftRate = shiftCount / distanceKm;
dwell = localGearDwell(t, shiftIdx);
huntingCount = localHuntingCount(gear, t, config.Thresholds.GearHuntingWindow_s);

driveEff = NaN(size(gbxOutPwr));
driveEff(driveMask) = gbxOutPwr(driveMask) ./ max(motorMechPwrIn(driveMask), eps);
regenEff = NaN(size(gbxOutPwr));
regenEff(regenMask) = abs(motorMechPwrIn(regenMask)) ./ max(abs(gbxOutPwr(regenMask)), eps);
lossShare = NaN(size(gbxOutPwr));
lossShare(driveMask) = 100 * max(gbxLoss(driveMask), 0) ./ max(motorMechPwrIn(driveMask), eps);

torqueTransferGain = NaN(size(gbxOutTrq));
validTransfer = isfinite(motorTrqIn) & isfinite(gbxOutTrq) & abs(motorTrqIn) > 1;
torqueTransferGain(validTransfer) = gbxOutTrq(validTransfer) ./ motorTrqIn(validTransfer);
ratioConsistency = NaN(size(gearRatio));
validRatioConsistency = isfinite(gearRatio) & isfinite(torqueTransferGain) & abs(gearRatio) > eps;
ratioConsistency(validRatioConsistency) = abs(torqueTransferGain(validRatioConsistency) - gearRatio(validRatioConsistency)) ./ max(abs(gearRatio(validRatioConsistency)), eps) * 100;

lossEnergy = RCA_TrapzFinite(t, max(gbxLoss, 0)) / 3600;
outDriveEnergy = RCA_TrapzFinite(t, max(gbxOutPwr, 0)) / 3600;
outRegenEnergy = RCA_TrapzFinite(t, max(-gbxOutPwr, 0)) / 3600;

shiftTorqueDipPct = localShiftTorqueDip(t, gbxOutTrq, shiftIdx);
shiftLossSpike_kW = localShiftLossSpike(t, gbxLoss, shiftIdx);

rows = RCA_AddKPI(rows, 'Gear Shift Count', shiftCount, 'count', ...
    'Operation', 'Transmission', 'gr_num', ...
    'Total number of observed actual-gear changes.');
rows = RCA_AddKPI(rows, 'Gear Shift Rate', shiftRate, 'shifts/km', ...
    'Operation', 'Transmission', 'gr_num + trip distance', ...
    'Shift density normalized by trip distance.');
rows = RCA_AddKPI(rows, 'Mean Gear Dwell', mean(dwell, 'omitnan'), 's', ...
    'Operation', 'Transmission', 'gr_num', ...
    'Average time between consecutive gear transitions.');
rows = RCA_AddKPI(rows, 'Gear Hunting Count', huntingCount, 'count', ...
    'Operation', 'Transmission', 'gr_num', ...
    'A-B-A reversals within the configured hunting window.');
rows = RCA_AddKPI(rows, 'Transmission Loss Energy', lossEnergy, 'kWh', ...
    'Losses', 'Transmission', 'gbx_pwr_loss', ...
    'Integrated positive gearbox loss power.');
rows = RCA_AddKPI(rows, 'Average Gearbox Loss Power', mean(max(gbxLoss, 0), 'omitnan'), 'kW', ...
    'Losses', 'Transmission', 'gbx_pwr_loss', ...
    'Mean positive gearbox loss power.');
rows = RCA_AddKPI(rows, 'Average Drive Transmission Efficiency', 100 * mean(driveEff, 'omitnan'), '%', ...
    'Efficiency', 'Transmission', 'input/output mechanical power', ...
    'Computed during positive mechanical power transfer.');
rows = RCA_AddKPI(rows, 'Average Regen Transmission Efficiency', 100 * mean(regenEff, 'omitnan'), '%', ...
    'Efficiency', 'Transmission', 'input/output mechanical power', ...
    'Computed during negative mechanical power transfer.');
rows = RCA_AddKPI(rows, 'High Loss Share in Drive', 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, driveMask), '%', ...
    'Losses', 'Transmission', 'gbx loss power / motor mechanical power', ...
    sprintf('Drive samples above %.1f%% loss share are flagged as high-loss gearbox operation.', config.Thresholds.HighLossShare_pct));
rows = RCA_AddKPI(rows, 'Output Drive Energy', outDriveEnergy, 'kWh', ...
    'Energy', 'Transmission', 'gbx_out_trq + gbx_out_spd', ...
    'Integrated positive gearbox output mechanical energy.');
rows = RCA_AddKPI(rows, 'Output Regen Energy', outRegenEnergy, 'kWh', ...
    'Energy', 'Transmission', 'gbx_out_trq + gbx_out_spd', ...
    'Integrated recovered mechanical energy magnitude at gearbox output.');
rows = RCA_AddKPI(rows, 'Mean Torque Transfer Gain', mean(torqueTransferGain(validTransfer), 'omitnan'), '-', ...
    'Transfer', 'Transmission', 'motor torque in + gearbox output torque', ...
    'Average gearbox torque multiplication or reduction factor.');
rows = RCA_AddKPI(rows, 'Gear Ratio Consistency Error', mean(ratioConsistency(validRatioConsistency), 'omitnan'), '%', ...
    'Transfer', 'Transmission', 'input/output torque ratio + gr_ratio', ...
    'Difference between effective torque-transfer gain and logged actual gear ratio.');
rows = RCA_AddKPI(rows, 'Mean Shift Torque Dip', mean(shiftTorqueDipPct, 'omitnan'), '%', ...
    'Shift Quality', 'Transmission', 'gbx_out_trq around shift events', ...
    'Average output torque dip through shift windows.');
rows = RCA_AddKPI(rows, 'Peak Shift Loss Spike', max(shiftLossSpike_kW, [], 'omitnan'), 'kW', ...
    'Shift Quality', 'Transmission', 'gbx_pwr_loss around shift events', ...
    'Largest gearbox loss spike observed around a shift.');

[uniqueGears, gearTimeShare, energyByGear, lossByGear, effByGear] = localPerGearStats(t, gear, gbxOutPwr, gbxLoss, driveEff);
for iGear = 1:numel(uniqueGears)
    gearLabel = sprintf('Gear %.0f', uniqueGears(iGear));
    rows = RCA_AddKPI(rows, [gearLabel ' Time Share'], gearTimeShare(iGear), '%', ...
        'Per Gear', 'Transmission', 'gr_num', ...
        'Share of total valid gear samples spent in this gear.');
    rows = RCA_AddKPI(rows, [gearLabel ' Output Drive Energy'], energyByGear(iGear), 'kWh', ...
        'Per Gear', 'Transmission', 'gr_num + gbx output power', ...
        'Positive gearbox output energy accumulated in this gear.');
    rows = RCA_AddKPI(rows, [gearLabel ' Loss Energy'], lossByGear(iGear), 'kWh', ...
        'Per Gear', 'Transmission', 'gr_num + gbx_pwr_loss', ...
        'Integrated gearbox loss energy in this gear.');
    rows = RCA_AddKPI(rows, [gearLabel ' Drive Efficiency'], effByGear(iGear), '%', ...
        'Per Gear', 'Transmission', 'gr_num + input/output mechanical power', ...
        'Average drive efficiency in this gear.');
end

summary(end + 1) = sprintf(['Transmission activity: %d gear shifts were observed, equivalent to %.1f shifts/km, with mean gear dwell of %.1f s.'], ...
    shiftCount, shiftRate, mean(dwell, 'omitnan'));
summary(end + 1) = sprintf(['Transmission energy path: %.2f kWh output drive energy and %.2f kWh gearbox loss energy were recorded. ', ...
    'Mean drive efficiency is %.1f%%.'], outDriveEnergy, lossEnergy, 100 * mean(driveEff, 'omitnan'));
summary(end + 1) = sprintf(['Shift quality: mean shift torque dip is %.1f%% and peak shift-loss spike is %.2f kW. ', ...
    'These indicate whether shifting interrupts wheel torque or drives transient gearbox losses.'], ...
    mean(shiftTorqueDipPct, 'omitnan'), max(shiftLossSpike_kW, [], 'omitnan'));
if any(validRatioConsistency)
    summary(end + 1) = sprintf(['Ratio consistency: mean difference between effective torque-transfer gain and logged gear ratio is %.1f%%. ', ...
        'Large persistent error can indicate sign-convention issues, model mismatch, or inconsistent logged ratio behavior.'], ...
        mean(ratioConsistency(validRatioConsistency), 'omitnan'));
end

if shiftRate > config.Thresholds.GearShiftRate_perkm
    recs(end + 1) = "Reduce shift activity by increasing ratio hysteresis or revisiting shift triggers; frequent shifting increases losses and can degrade drivability.";
    evidence(end + 1) = sprintf('Shift rate is %.1f shifts/km versus the %.1f shifts/km heuristic.', shiftRate, config.Thresholds.GearShiftRate_perkm);
end
if mean(dwell, 'omitnan') < config.Thresholds.MinGearDwell_s
    recs(end + 1) = "Increase minimum dwell or strengthen hold logic so the transmission does not change gear before torque and speed settle.";
    evidence(end + 1) = sprintf('Mean gear dwell is %.1f s versus the %.1f s heuristic.', mean(dwell, 'omitnan'), config.Thresholds.MinGearDwell_s);
end
if huntingCount > 0
    recs(end + 1) = "Review gear hunting around load and slope transitions; repeated reversals waste energy and destabilize torque delivery.";
    evidence(end + 1) = sprintf('%d gear-hunting events were detected.', huntingCount);
end
if 100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, driveMask) > 15
    recs(end + 1) = "Investigate gearbox efficiency degradation at the observed operating points; loss intensity is elevated for a material share of drive operation.";
    evidence(end + 1) = sprintf('High-loss drive share is %.1f%% above the %.1f%% threshold.', ...
        100 * RCA_FractionTrue(lossShare >= config.Thresholds.HighLossShare_pct, driveMask), config.Thresholds.HighLossShare_pct);
end
if mean(shiftTorqueDipPct, 'omitnan') > 15
    recs(end + 1) = "Review shift torque-handshake logic between electric drive and transmission; output torque drops materially during shifts.";
    evidence(end + 1) = sprintf('Mean shift torque dip is %.1f%%.', mean(shiftTorqueDipPct, 'omitnan'));
end
if any(validRatioConsistency) && mean(ratioConsistency(validRatioConsistency), 'omitnan') > 20
    recs(end + 1) = "Check the consistency between actual gear ratio logging and gearbox torque transfer. Large mismatch can hide a model or signal-sign issue.";
    evidence(end + 1) = sprintf('Mean ratio consistency error is %.1f%%.', mean(ratioConsistency(validRatioConsistency), 'omitnan'));
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'Transmission');
plotFiles = localAppendPlotFile(plotFiles, localPlotOverview(figureFolder, t, gear, gearRatio, gbxOutTrq, gbxOutSpd, gbxLoss, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotGearUsage(figureFolder, vehSpeed, gear, uniqueGears, energyByGear, lossByGear, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotTransferAndLoss(figureFolder, torqueTransferGain, gearRatio, lossShare, driveEff, regenEff, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotShiftQuality(figureFolder, t, shiftIdx, shiftTorqueDipPct, shiftLossSpike_kW, gear, gbxOutTrq, gbxLoss, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Transmission", recs, evidence);
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

function shiftEvents = localShiftEvents(gear)
shiftEvents = false(size(gear));
if numel(gear) < 2
    return;
end
validStep = isfinite(gear(2:end)) & isfinite(gear(1:end-1));
shiftEvents(2:end) = validStep & abs(diff(gear)) > 0.05;
end

function shiftMask = localShiftMask(gear)
shiftMask = localShiftEvents(gear);
if any(shiftMask)
    shiftMask = shiftMask | [false; shiftMask(1:end-1)] | [shiftMask(2:end); false];
end
end

function dwell = localGearDwell(t, shiftIdx)
if isempty(t)
    dwell = NaN;
    return;
end
bounds = [1; shiftIdx(:); numel(t)];
if numel(bounds) < 2
    dwell = t(end) - t(1);
else
    dwell = diff(t(bounds));
end
end

function huntingCount = localHuntingCount(gear, t, huntingWindow_s)
huntingCount = 0;
shiftIdx = find(localShiftEvents(gear));
for iShift = 3:numel(shiftIdx)
    if gear(shiftIdx(iShift)) == gear(shiftIdx(iShift - 2)) && ...
            (t(shiftIdx(iShift)) - t(shiftIdx(iShift - 2))) <= huntingWindow_s
        huntingCount = huntingCount + 1;
    end
end
end

function dipPct = localShiftTorqueDip(t, gbxOutTrq, shiftIdx)
dipPct = NaN(numel(shiftIdx), 1);
for iShift = 1:numel(shiftIdx)
    idx = shiftIdx(iShift);
    mask = localWindowMask(t, t(idx), 1.0, 1.0);
    if ~any(mask)
        continue;
    end
    windowTorque = abs(gbxOutTrq(mask));
    preMask = mask & t < t(idx);
    postMask = mask & t > t(idx);
    refLevel = mean([mean(abs(gbxOutTrq(preMask)), 'omitnan'), mean(abs(gbxOutTrq(postMask)), 'omitnan')], 'omitnan');
    if isfinite(refLevel) && refLevel > eps
        dipPct(iShift) = 100 * max(refLevel - min(windowTorque, [], 'omitnan'), 0) / refLevel;
    end
end
end

function spike_kW = localShiftLossSpike(t, gbxLoss, shiftIdx)
spike_kW = NaN(numel(shiftIdx), 1);
for iShift = 1:numel(shiftIdx)
    idx = shiftIdx(iShift);
    preMask = localWindowMask(t, t(idx), 1.0, 0.0);
    postMask = localWindowMask(t, t(idx), 0.0, 1.0);
    baseline = mean(max(gbxLoss(preMask), 0), 'omitnan');
    localPeak = max(max(gbxLoss(postMask), 0), [], 'omitnan');
    spike_kW(iShift) = localPeak - baseline;
end
end

function mask = localWindowMask(t, centerTime, before_s, after_s)
mask = isfinite(t) & t >= (centerTime - before_s) & t <= (centerTime + after_s);
end

function [uniqueGears, timeShare, driveEnergy, lossEnergy, effPct] = localPerGearStats(t, gear, gbxOutPwr, gbxLoss, driveEff)
validGear = isfinite(gear);
uniqueGears = unique(round(gear(validGear)));
uniqueGears = uniqueGears(:)';
timeShare = NaN(size(uniqueGears));
driveEnergy = NaN(size(uniqueGears));
lossEnergy = NaN(size(uniqueGears));
effPct = NaN(size(uniqueGears));
for iGear = 1:numel(uniqueGears)
    gearMask = validGear & round(gear) == uniqueGears(iGear);
    timeShare(iGear) = 100 * RCA_FractionTrue(gearMask, validGear);
    driveEnergy(iGear) = RCA_TrapzFinite(t(gearMask), max(gbxOutPwr(gearMask), 0)) / 3600;
    lossEnergy(iGear) = RCA_TrapzFinite(t(gearMask), max(gbxLoss(gearMask), 0)) / 3600;
    effPct(iGear) = 100 * mean(driveEff(gearMask), 'omitnan');
end
end

function plotFile = localPlotOverview(outputFolder, t, gear, gearRatio, gbxOutTrq, gbxOutSpd, gbxLoss, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(4, 1, 1);
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
title('Actual Gear Number');
ylabel('Gear (-)');
grid on;

subplot(4, 1, 2);
plot(t, gearRatio, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Actual Gear Ratio');
ylabel('Ratio (-)');
grid on;

subplot(4, 1, 3);
plot(t, gbxOutTrq, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
yyaxis right;
plot(t, gbxOutSpd, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
title('Gearbox Output Torque and Speed');
xlabel('Time (s)');
yyaxis left;
ylabel('Torque (Nm)');
yyaxis right;
ylabel('Speed (rad/s)');
grid on;

subplot(4, 1, 4);
plot(t, gbxLoss, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Gearbox Loss Power');
xlabel('Time (s)');
ylabel('Loss power (kW)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Transmission_Overview', config));
close(fig);
end

function plotFile = localPlotGearUsage(outputFolder, vehSpeed, gear, uniqueGears, energyByGear, lossByGear, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid = isfinite(vehSpeed) & isfinite(gear);
scatter(vehSpeed(valid), gear(valid), 12, gear(valid), 'filled');
title('Gear Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Gear (-)');
grid on;

subplot(2, 2, 2);
histogram(gear(isfinite(gear)), 'FaceColor', config.Plot.Colors.Gear);
title('Gear Usage Histogram');
xlabel('Gear (-)');
ylabel('Samples');
grid on;

subplot(2, 2, 3);
bar(uniqueGears, energyByGear, 'FaceColor', config.Plot.Colors.Vehicle);
title('Output Drive Energy by Gear');
xlabel('Gear (-)');
ylabel('Energy (kWh)');
grid on;

subplot(2, 2, 4);
bar(uniqueGears, lossByGear, 'FaceColor', config.Plot.Colors.Warning);
title('Loss Energy by Gear');
xlabel('Gear (-)');
ylabel('Loss energy (kWh)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Transmission_GearUsage', config));
close(fig);
end

function plotFile = localPlotTransferAndLoss(outputFolder, torqueTransferGain, gearRatio, lossShare, driveEff, regenEff, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
validRatio = isfinite(gearRatio) & isfinite(torqueTransferGain);
scatter(gearRatio(validRatio), torqueTransferGain(validRatio), 12, config.Plot.Colors.Vehicle, 'filled');
title('Effective Torque Gain Versus Logged Gear Ratio');
xlabel('Logged gear ratio (-)');
ylabel('Effective torque transfer gain (-)');
grid on;

subplot(2, 2, 2);
validLoss = isfinite(lossShare) & isfinite(driveEff);
scatter(lossShare(validLoss), driveEff(validLoss) * 100, 12, config.Plot.Colors.Warning, 'filled');
title('Drive Efficiency Versus Loss Share');
xlabel('Loss share (%)');
ylabel('Drive efficiency (%)');
grid on;

subplot(2, 2, 3);
histogram(driveEff(isfinite(driveEff)) * 100, 'FaceColor', config.Plot.Colors.Motor);
title('Drive Efficiency Distribution');
xlabel('Efficiency (%)');
ylabel('Samples');
grid on;

subplot(2, 2, 4);
histogram(regenEff(isfinite(regenEff)) * 100, 'FaceColor', config.Plot.Colors.Battery);
title('Regen Efficiency Distribution');
xlabel('Efficiency (%)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Transmission_TransferAndLoss', config));
close(fig);
end

function plotFile = localPlotShiftQuality(outputFolder, t, shiftIdx, shiftTorqueDipPct, shiftLossSpike_kW, gear, gbxOutTrq, gbxLoss, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
if isempty(shiftIdx)
    text(0.5, 0.5, 'No shift events detected', 'HorizontalAlignment', 'center');
    axis off;
else
    bar(1:numel(shiftIdx), shiftTorqueDipPct, 'FaceColor', config.Plot.Colors.Vehicle);
    title('Torque Dip by Shift Event');
    xlabel('Shift event index');
    ylabel('Dip (%)');
    grid on;
end

subplot(2, 2, 2);
if isempty(shiftIdx)
    text(0.5, 0.5, 'No shift events detected', 'HorizontalAlignment', 'center');
    axis off;
else
    bar(1:numel(shiftIdx), shiftLossSpike_kW, 'FaceColor', config.Plot.Colors.Warning);
    title('Loss Spike by Shift Event');
    xlabel('Shift event index');
    ylabel('Loss spike (kW)');
    grid on;
end

subplot(2, 1, 2);
plot(t, gbxOutTrq, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, gbxLoss, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
for iShift = 1:numel(shiftIdx)
    xline(t(shiftIdx(iShift)), ':', 'Color', config.Plot.Colors.Gear);
end
yyaxis right;
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', 1.0);
title('Shift Context on Output Torque and Loss');
xlabel('Time (s)');
yyaxis left;
ylabel('Torque / loss');
yyaxis right;
ylabel('Gear (-)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Transmission_ShiftQuality', config));
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
