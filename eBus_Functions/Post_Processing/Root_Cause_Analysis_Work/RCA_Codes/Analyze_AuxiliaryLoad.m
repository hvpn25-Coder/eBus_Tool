function result = Analyze_AuxiliaryLoad(analysisData, outputPaths, config)
% Analyze_AuxiliaryLoad  Auxiliary electrical-load RCA using current, voltage, and derived power burden.

result = localInitResult("AUXILIARY LOAD", {'aux_curr', 'aux_volt'}, {'batt_pwr', 'veh_vel', 'batt_soc'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Auxiliary load analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Auxiliary Load", recs, evidence);
    result.SummaryText = summary;
    return;
end

auxCurr = d.auxiliaryCurrent_A(:);
auxVolt = d.auxiliaryVoltage_V(:);
auxPwr = d.auxiliaryPower_kW(:);
battPwr = d.batteryPower_kW(:);
vehSpeed = d.vehVel_kmh(:);
battSoc = d.batterySOC_pct(:);

if ~any(isfinite(auxCurr))
    auxCurr = localAlignedSignal(analysisData.Signals, 'aux_curr', n);
end
if ~any(isfinite(auxVolt))
    auxVolt = localAlignedSignal(analysisData.Signals, 'aux_volt', n);
end
if ~any(isfinite(auxPwr))
    auxPwr = auxCurr .* auxVolt / 1000;
end

if all(isnan(auxPwr))
    result.Warnings(end + 1) = "Auxiliary load analysis skipped because auxiliary current/voltage are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Auxiliary Load", recs, evidence);
    result.SummaryText = summary;
    return;
end

movingMask = isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh;
stationaryMask = isfinite(vehSpeed) & vehSpeed <= config.Thresholds.StopSpeed_kmh;
highAuxMask = isfinite(auxPwr) & auxPwr >= config.Thresholds.HighAuxPower_kW;
activeDischargeMask = isfinite(auxPwr) & isfinite(battPwr) & battPwr > 0;

auxEnergy = RCA_TrapzFinite(t, max(auxPwr, 0)) / 3600;
movingAuxEnergy = RCA_TrapzFinite(t(movingMask), max(auxPwr(movingMask), 0)) / 3600;
stationaryAuxEnergy = RCA_TrapzFinite(t(stationaryMask), max(auxPwr(stationaryMask), 0)) / 3600;
netBattEnergy = RCA_TrapzFinite(t, max(battPwr, 0)) / 3600;

auxShare = NaN;
if netBattEnergy > 0
    auxShare = 100 * auxEnergy / netBattEnergy;
end

stationaryShare = 100 * stationaryAuxEnergy / max(auxEnergy, eps);
movingShare = 100 * movingAuxEnergy / max(auxEnergy, eps);
rangePenalty = d.tripDistance_km * auxShare / 100;

instantShare = NaN(size(auxPwr));
instantShare(activeDischargeMask) = 100 * auxPwr(activeDischargeMask) ./ max(battPwr(activeDischargeMask), eps);

rows = RCA_AddKPI(rows, 'Auxiliary Energy', auxEnergy, 'kWh', ...
    'Energy', 'Auxiliary Load', 'aux_curr + aux_volt', ...
    'Derived from auxiliary current and voltage.');
rows = RCA_AddKPI(rows, 'Average Auxiliary Power', mean(auxPwr, 'omitnan'), 'kW', ...
    'Efficiency', 'Auxiliary Load', 'aux_curr + aux_volt', ...
    'Average auxiliary power over the full trip.');
rows = RCA_AddKPI(rows, 'Peak Auxiliary Power', max(auxPwr, [], 'omitnan'), 'kW', ...
    'Efficiency', 'Auxiliary Load', 'aux_curr + aux_volt', ...
    'Maximum observed auxiliary power.');
rows = RCA_AddKPI(rows, 'Average Auxiliary Current', mean(auxCurr, 'omitnan'), 'A', ...
    'Operation', 'Auxiliary Load', 'aux_curr', ...
    'Average auxiliary current over the full trip.');
rows = RCA_AddKPI(rows, 'Average Auxiliary Voltage', mean(auxVolt, 'omitnan'), 'V', ...
    'Operation', 'Auxiliary Load', 'aux_volt', ...
    'Average auxiliary voltage over the full trip.');
rows = RCA_AddKPI(rows, 'Auxiliary Energy Share of Battery Discharge', auxShare, '%', ...
    'Efficiency', 'Auxiliary Load', 'auxiliary power + battery power', ...
    'Share of discharge-positive battery energy attributed to auxiliaries.');
rows = RCA_AddKPI(rows, 'Stationary Auxiliary Energy Share', stationaryShare, '%', ...
    'Operation', 'Auxiliary Load', 'auxiliary power + vehicle speed', ...
    'Fraction of auxiliary energy consumed while vehicle speed is near zero.');
rows = RCA_AddKPI(rows, 'Moving Auxiliary Energy Share', movingShare, '%', ...
    'Operation', 'Auxiliary Load', 'auxiliary power + vehicle speed', ...
    'Fraction of auxiliary energy consumed while the vehicle is moving.');
rows = RCA_AddKPI(rows, 'High Auxiliary Power Share', 100 * RCA_FractionTrue(highAuxMask, isfinite(auxPwr)), '%', ...
    'Efficiency', 'Auxiliary Load', 'auxiliary power', ...
    sprintf('Share of samples above the %.1f kW high-auxiliary heuristic.', config.Thresholds.HighAuxPower_kW));
rows = RCA_AddKPI(rows, 'High Auxiliary Power Share While Stationary', 100 * RCA_FractionTrue(highAuxMask, stationaryMask), '%', ...
    'Operation', 'Auxiliary Load', 'auxiliary power + vehicle speed', ...
    'Share of stationary samples with high auxiliary power.');
rows = RCA_AddKPI(rows, 'High Auxiliary Power Share While Moving', 100 * RCA_FractionTrue(highAuxMask, movingMask), '%', ...
    'Operation', 'Auxiliary Load', 'auxiliary power + vehicle speed', ...
    'Share of moving samples with high auxiliary power.');
rows = RCA_AddKPI(rows, 'Average Instantaneous Auxiliary Share', mean(instantShare, 'omitnan'), '%', ...
    'Efficiency', 'Auxiliary Load', 'auxiliary power + battery power', ...
    'Average instantaneous auxiliary share of battery discharge power.');
rows = RCA_AddKPI(rows, 'Approximate Distance Penalty from Auxiliary Share', rangePenalty, 'km', ...
    'Range', 'Auxiliary Load', 'trip distance + auxiliary share', ...
    'Approximation assuming overall trip behavior remains similar.');
rows = RCA_AddKPI(rows, 'Low-SoC High Auxiliary Share', 100 * RCA_FractionTrue(highAuxMask, isfinite(auxPwr) & battSoc <= config.Thresholds.LowSOC_pct), '%', ...
    'Context', 'Auxiliary Load', 'auxiliary power + batt_soc', ...
    'Share of low-SoC samples with high auxiliary power.');

summary(end + 1) = sprintf(['Auxiliaries consume %.2f kWh, equal to %.1f%% of discharge-positive battery energy, with %.1f kW average power and %.1f kW peak power.'], ...
    auxEnergy, auxShare, mean(auxPwr, 'omitnan'), max(auxPwr, [], 'omitnan'));
summary(end + 1) = sprintf(['Usage context: %.1f%% of auxiliary energy is consumed while stationary and %.1f%% while moving. ', ...
    'This helps identify idle loads versus route-coupled auxiliary burden.'], stationaryShare, movingShare);
summary(end + 1) = sprintf(['Auxiliary burden context: high-auxiliary-power share is %.1f%% overall, with %.1f%% while stationary.'], ...
    100 * RCA_FractionTrue(highAuxMask, isfinite(auxPwr)), 100 * RCA_FractionTrue(highAuxMask, stationaryMask));

if mean(auxPwr, 'omitnan') > config.Thresholds.HighAuxPower_kW
    recs(end + 1) = "Reduce continuous auxiliary demand or defer peaks from propulsion-critical windows; auxiliaries are a first-order efficiency driver.";
    evidence(end + 1) = sprintf('Average auxiliary power is %.1f kW against the %.1f kW heuristic threshold.', ...
        mean(auxPwr, 'omitnan'), config.Thresholds.HighAuxPower_kW);
end
if stationaryShare > 30
    recs(end + 1) = "A large stationary auxiliary burden suggests HVAC, compressor, or accessory logic deserves a dedicated idle-energy review.";
    evidence(end + 1) = sprintf('%.1f%% of auxiliary energy is consumed while the vehicle is effectively stationary.', stationaryShare);
end
if auxShare > config.Thresholds.HighAuxShare_pct
    recs(end + 1) = "Auxiliary share of discharge energy is high enough to materially affect range. Review auxiliary scheduling, duty-cycling, and DC-link loading.";
    evidence(end + 1) = sprintf('Auxiliary energy share is %.1f%% versus the %.1f%% heuristic threshold.', auxShare, config.Thresholds.HighAuxShare_pct);
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'AuxiliaryLoad');
plotFiles = localAppendPlotFile(plotFiles, localPlotOverview(figureFolder, t, auxCurr, auxVolt, auxPwr, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotBatteryShare(figureFolder, t, auxPwr, instantShare, battPwr, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotContext(figureFolder, vehSpeed, battSoc, auxPwr, auxCurr, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotEnergyBreakdown(figureFolder, auxEnergy, stationaryAuxEnergy, movingAuxEnergy, netBattEnergy, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Auxiliary Load", recs, evidence);
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

function plotFile = localPlotOverview(outputFolder, t, auxCurr, auxVolt, auxPwr, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, auxCurr, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
title('Auxiliary Current');
ylabel('Current (A)');
grid on;

subplot(3, 1, 2);
plot(t, auxVolt, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Auxiliary Voltage');
ylabel('Voltage (V)');
grid on;

subplot(3, 1, 3);
plot(t, auxPwr, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.HighAuxPower_kW, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Auxiliary Power');
xlabel('Time (s)');
ylabel('Power (kW)');
legend({'Auxiliary power', 'High auxiliary threshold'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'AuxiliaryLoad_Overview', config));
close(fig);
end

function plotFile = localPlotBatteryShare(outputFolder, t, auxPwr, instantShare, battPwr, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 1, 1);
plot(t, auxPwr, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(battPwr, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Auxiliary Power Versus Battery Discharge Power');
ylabel('Power (kW)');
legend({'Auxiliary power', 'Battery discharge power'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
plot(t, instantShare, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.HighAuxShare_pct, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Auxiliary Share of Instantaneous Battery Discharge Power');
xlabel('Time (s)');
ylabel('Share (%)');
legend({'Instantaneous auxiliary share', 'High auxiliary-share threshold'}, 'Location', 'best');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'AuxiliaryLoad_BatteryShare', config));
close(fig);
end

function plotFile = localPlotContext(outputFolder, vehSpeed, battSoc, auxPwr, auxCurr, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid1 = isfinite(vehSpeed) & isfinite(auxPwr);
scatter(vehSpeed(valid1), auxPwr(valid1), 12, config.Plot.Colors.Auxiliary, 'filled');
title('Auxiliary Power Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 2);
valid2 = isfinite(battSoc) & isfinite(auxPwr);
scatter(battSoc(valid2), auxPwr(valid2), 12, config.Plot.Colors.Battery, 'filled');
title('Auxiliary Power Versus SoC');
xlabel('SoC (%)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 3);
valid3 = isfinite(auxCurr) & isfinite(auxPwr);
scatter(auxCurr(valid3), auxPwr(valid3), 12, config.Plot.Colors.Vehicle, 'filled');
title('Auxiliary Power Versus Current');
xlabel('Current (A)');
ylabel('Power (kW)');
grid on;

subplot(2, 2, 4);
histogram(auxPwr(isfinite(auxPwr)), 40, 'FaceColor', config.Plot.Colors.Auxiliary);
title('Auxiliary Power Distribution');
xlabel('Power (kW)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'AuxiliaryLoad_Context', config));
close(fig);
end

function plotFile = localPlotEnergyBreakdown(outputFolder, auxEnergy, stationaryAuxEnergy, movingAuxEnergy, netBattEnergy, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

bar(categorical({'Aux total', 'Aux stationary', 'Aux moving', 'Battery discharge'}), ...
    [auxEnergy, stationaryAuxEnergy, movingAuxEnergy, netBattEnergy], 'FaceColor', config.Plot.Colors.Vehicle);
title('Auxiliary Energy Breakdown');
ylabel('Energy (kWh)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'AuxiliaryLoad_EnergyBreakdown', config));
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
