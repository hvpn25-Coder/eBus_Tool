function result = Analyze_Environment(analysisData, outputPaths, config)
% Analyze_Environment  Environmental severity analysis.

result = localInitResult("ENVIRONMENT", {'road_slp'}, {'veh_des_vel', 'amb_temp', 'veh_vel'});
t = analysisData.Derived.time_s;
distanceStep = analysisData.Derived.distanceStep_km;

roadSlope = RCA_GetSignalData(analysisData.Signals, 'road_slp');
desiredSpeed = RCA_GetSignalData(analysisData.Signals, 'veh_des_vel');
ambientTemp = RCA_GetSignalData(analysisData.Signals, 'amb_temp');
vehicleSpeed = analysisData.Derived.vehVel_kmh;

rows = cell(0, 7);
summary = strings(0, 1);

if roadSlope.Available
    slope = analysisData.Derived.roadSlope_pct;
    uphillShare = 100 * sum(distanceStep(slope > config.Thresholds.UphillSlope_pct), 'omitnan') / max(sum(distanceStep, 'omitnan'), eps);
    steepShare = 100 * sum(distanceStep(slope > config.Thresholds.SteepSlope_pct), 'omitnan') / max(sum(distanceStep, 'omitnan'), eps);
    rows = RCA_AddKPI(rows, 'Mean Road Slope', mean(slope, 'omitnan'), '%', 'Operation', 'Environment', 'road_slp', 'Complete if road_slp is available.');
    rows = RCA_AddKPI(rows, 'Maximum Road Slope', max(slope, [], 'omitnan'), '%', 'Operation', 'Environment', 'road_slp', 'Complete if road_slp is available.');
    rows = RCA_AddKPI(rows, 'Uphill Distance Share', uphillShare, '%', 'Range', 'Environment', 'road_slp + veh_pos or veh_vel', 'Derived from available trip distance basis.');
    rows = RCA_AddKPI(rows, 'Steep Uphill Distance Share', steepShare, '%', 'Range', 'Environment', 'road_slp + veh_pos or veh_vel', 'Severe uphill threshold comes from RCA_Config.');
    summary(end + 1) = sprintf('Environment severity: uphill distance share is %.1f%% and steep uphill share is %.1f%%.', uphillShare, steepShare);
end

if desiredSpeed.Available
    rows = RCA_AddKPI(rows, 'Average Desired Speed', mean(analysisData.Derived.speedDemand_kmh, 'omitnan'), 'km/h', 'Performance', 'Environment', 'veh_des_vel', 'Demand trace available.');
    rows = RCA_AddKPI(rows, 'Peak Desired Speed', max(analysisData.Derived.speedDemand_kmh, [], 'omitnan'), 'km/h', 'Performance', 'Environment', 'veh_des_vel', 'Demand trace available.');
end

if ambientTemp.Available
    amb = analysisData.Derived.ambientTemp_C;
    rows = RCA_AddKPI(rows, 'Mean Ambient Temperature', mean(amb, 'omitnan'), 'degC', 'Operation', 'Environment', 'amb_temp', 'Complete if ambient signal is available.');
    rows = RCA_AddKPI(rows, 'Ambient Temperature Range', max(amb, [], 'omitnan') - min(amb, [], 'omitnan'), 'degC', 'Operation', 'Environment', 'amb_temp', 'Complete if ambient signal is available.');
    if mean(amb, 'omitnan') > 35 || mean(amb, 'omitnan') < 5
        summary(end + 1) = "Ambient conditions are thermally severe enough to influence battery and auxiliary behaviour.";
    end
end

recs = strings(0, 1);
evidence = strings(0, 1);
if roadSlope.Available && any(analysisData.Derived.roadSlope_pct > config.Thresholds.SteepSlope_pct)
    recs(end + 1) = "Keep grade severity visible in future simulation reviews; uphill route severity materially changes range and performance conclusions.";
    evidence(end + 1) = "Steep slope share exceeded the configured severe-grade threshold.";
end
if ambientTemp.Available && (mean(analysisData.Derived.ambientTemp_C, 'omitnan') > 35 || mean(analysisData.Derived.ambientTemp_C, 'omitnan') < 5)
    recs(end + 1) = "Correlate ambient temperature with battery temperature and auxiliary power to isolate thermal penalties more confidently.";
    evidence(end + 1) = "Ambient temperature is outside a mild operating band.";
end

if any([roadSlope.Available, desiredSpeed.Available, ambientTemp.Available])
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(3, 1, 1);
    plot(t, vehicleSpeed, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    if desiredSpeed.Available
        plot(t, analysisData.Derived.speedDemand_kmh, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
        legend({'Vehicle speed', 'Desired speed'}, 'Location', 'best');
    end
    title('Environment Context: Speed Demand and Response');
    ylabel('Speed (km/h)');
    grid on;

    subplot(3, 1, 2);
    if roadSlope.Available
        plot(t, analysisData.Derived.roadSlope_pct, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
        ylabel('Slope (%)');
    else
        text(0.1, 0.5, 'Road slope unavailable', 'Units', 'normalized');
    end
    title('Route Grade');
    grid on;

    subplot(3, 1, 3);
    if ambientTemp.Available
        plot(t, analysisData.Derived.ambientTemp_C, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
        ylabel('Ambient (degC)');
    else
        text(0.1, 0.5, 'Ambient temperature unavailable', 'Units', 'normalized');
    end
    xlabel('Time (s)');
    title('Ambient Temperature');
    grid on;

    result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Environment'), 'Environment_Overview', config));
    close(fig);
end

result.Available = ~isempty(rows);
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Environment", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
