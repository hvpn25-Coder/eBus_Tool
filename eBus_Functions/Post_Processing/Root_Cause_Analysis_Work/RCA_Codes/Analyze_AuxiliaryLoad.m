function result = Analyze_AuxiliaryLoad(analysisData, outputPaths, config)
% Analyze_AuxiliaryLoad  Auxiliary power burden and range penalty analysis.

result = localInitResult("AUXILIARY LOAD", {'aux_curr', 'aux_volt'}, {'batt_pwr', 'veh_vel'});
t = analysisData.Derived.time_s;
auxPwr = analysisData.Derived.auxiliaryPower_kW;
battPwr = analysisData.Derived.batteryPower_kW;
vehSpeed = analysisData.Derived.vehVel_kmh;

rows = cell(0, 7);
summary = strings(0, 1);
if all(isnan(auxPwr))
    result.Warnings(end + 1) = "Auxiliary load signal is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Auxiliary Load", strings(0, 1), strings(0, 1));
    return;
end

auxEnergy = trapz(t, max(auxPwr, 0)) / 3600;
netBattEnergy = trapz(t, max(battPwr, 0)) / 3600;
auxShare = 100 * auxEnergy / max(netBattEnergy, eps);
stationaryMask = vehSpeed <= config.Thresholds.StopSpeed_kmh;
stationaryShare = 100 * trapz(t(stationaryMask), max(auxPwr(stationaryMask), 0)) / 3600 / max(auxEnergy, eps);
rangePenalty = analysisData.Derived.tripDistance_km * auxShare / 100;

rows = RCA_AddKPI(rows, 'Auxiliary Energy', auxEnergy, 'kWh', 'Energy', 'Auxiliary Load', 'aux_curr + aux_volt', 'Derived from auxiliary current and voltage.');
rows = RCA_AddKPI(rows, 'Average Auxiliary Power', mean(auxPwr, 'omitnan'), 'kW', 'Efficiency', 'Auxiliary Load', 'aux_curr + aux_volt', 'Average over the full trip.');
rows = RCA_AddKPI(rows, 'Auxiliary Energy Share', auxShare, '%', 'Efficiency', 'Auxiliary Load', 'auxiliary power + battery power', 'Share of positive battery energy attributed to auxiliaries.');
rows = RCA_AddKPI(rows, 'Stationary Auxiliary Energy Share', stationaryShare, '%', 'Operation', 'Auxiliary Load', 'auxiliary power + vehicle speed', 'Fraction of auxiliary energy consumed while vehicle speed is near zero.');
rows = RCA_AddKPI(rows, 'Approximate Distance Penalty from Auxiliary Share', rangePenalty, 'km', 'Range', 'Auxiliary Load', 'trip distance + auxiliary share', 'Approximation assumes overall trip behaviour remains similar.');
summary(end + 1) = sprintf('Auxiliaries consume %.2f kWh, equal to %.1f%% of positive battery energy.', auxEnergy, auxShare);

recs = strings(0, 1);
evidence = strings(0, 1);
if mean(auxPwr, 'omitnan') > config.Thresholds.HighAuxPower_kW
    recs(end + 1) = "Reduce continuous auxiliary demand or defer peaks from propulsion-critical windows; auxiliaries are a first-order efficiency driver.";
    evidence(end + 1) = sprintf('Average auxiliary power is %.1f kW against a %.1f kW heuristic threshold.', ...
        mean(auxPwr, 'omitnan'), config.Thresholds.HighAuxPower_kW);
end
if stationaryShare > 30
    recs(end + 1) = "A large stationary auxiliary burden suggests HVAC, compressor, or accessory logic deserves a dedicated idle-energy review.";
    evidence(end + 1) = sprintf('%.1f%% of auxiliary energy is consumed while the vehicle is effectively stationary.', stationaryShare);
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, auxPwr, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
title('Auxiliary Power');
ylabel('Power (kW)');
grid on;

subplot(2, 1, 2);
plot(t, 100 * auxPwr ./ max(max(battPwr, 0), eps), 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Auxiliary Share of Instantaneous Positive Battery Power');
xlabel('Time (s)');
ylabel('Share (%)');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'AuxiliaryLoad'), 'AuxiliaryLoad_Overview', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Auxiliary Load", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
