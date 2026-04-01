function result = Analyze_Environment(analysisData, outputPaths, config)
% Analyze_Environment  Route severity, duty-cycle demand, and thermal context analysis.

result = localInitResult("ENVIRONMENT", {'road_slp'}, {'veh_des_vel', 'amb_temp', 'veh_vel'});

derived = analysisData.Derived;
t = derived.time_s;
dt = derived.dt_s;

roadSlopeSignal = RCA_GetSignalData(analysisData.Signals, 'road_slp');
desiredSpeedSignal = RCA_GetSignalData(analysisData.Signals, 'veh_des_vel');
ambientTempSignal = RCA_GetSignalData(analysisData.Signals, 'amb_temp');

desiredSpeed = derived.speedDemand_kmh;
actualSpeed = derived.vehVel_kmh;
roadSlope = derived.roadSlope_pct;
ambientTemp = derived.ambientTemp_C;

rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if ~any([roadSlopeSignal.Available, desiredSpeedSignal.Available, ambientTempSignal.Available])
    result.Warnings(end + 1) = "Environment subsystem signals are unavailable.";
    result.KPITable = localRemoveSubsystemColumn(RCA_FinalizeKPITable(rows));
    result.Suggestions = RCA_MakeSuggestionTable("Environment", strings(0, 1), strings(0, 1));
    return;
end

speedForContext = desiredSpeed;
speedBasisLabel = "veh_des_vel";
if ~any(isfinite(speedForContext)) && any(isfinite(actualSpeed))
    speedForContext = actualSpeed;
    speedBasisLabel = "veh_vel";
end

[distanceStep_km, cumulativeDistance_km, distanceBasisLabel, distanceBasisNote] = localDistanceBasis(derived, speedForContext, dt, config);
routeDistance_km = sum(max(distanceStep_km, 0), 'omitnan');
if ~isfinite(routeDistance_km) || routeDistance_km <= 0
    routeDistance_km = NaN;
end

if desiredSpeedSignal.Available && any(isfinite(desiredSpeed))
    desiredAccel = localDifferentiateTrace(t, desiredSpeed / 3.6);
    validDesired = isfinite(desiredSpeed);
    validDesiredAccel = isfinite(desiredAccel);

    stopShare = 100 * RCA_FractionTrue(desiredSpeed <= config.Thresholds.StopSpeed_kmh, validDesired);
    creepShare = 100 * RCA_FractionTrue(desiredSpeed > config.Thresholds.StopSpeed_kmh & desiredSpeed <= config.Thresholds.CreepSpeed_kmh, validDesired);
    urbanShare = 100 * RCA_FractionTrue(desiredSpeed > config.Thresholds.CreepSpeed_kmh & desiredSpeed < config.Thresholds.UrbanSpeedUpper_kmh, validDesired);
    mediumSpeedShare = 100 * RCA_FractionTrue(desiredSpeed >= config.Thresholds.UrbanSpeedUpper_kmh & desiredSpeed < config.Thresholds.HighwaySpeed_kmh, validDesired);
    highwayShare = 100 * RCA_FractionTrue(desiredSpeed >= config.Thresholds.HighwaySpeed_kmh, validDesired);
    aggressiveShare = 100 * RCA_FractionTrue(abs(desiredAccel) >= config.Thresholds.SignificantAccel_mps2, validDesiredAccel);
    plannedStopCount = localCountEpisodes(desiredSpeed <= config.Thresholds.StopSpeed_kmh, t, config.Thresholds.MinStopDuration_s);
    positiveAccel95 = RCA_Percentile(desiredAccel(desiredAccel > 0), 95);
    negativeAccel95 = RCA_Percentile(-desiredAccel(desiredAccel < 0), 95);

    rows = RCA_AddKPI(rows, 'Average Desired Speed', mean(desiredSpeed, 'omitnan'), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', 'Average demanded vehicle speed from the environment subsystem.');
    rows = RCA_AddKPI(rows, 'Peak Desired Speed', max(desiredSpeed, [], 'omitnan'), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', 'Peak demanded speed.');
    rows = RCA_AddKPI(rows, 'Desired Speed 95th Percentile', RCA_Percentile(desiredSpeed, 95), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', 'Upper-end speed demand severity.');
    rows = RCA_AddKPI(rows, 'Demand Stop Time Share', stopShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand-cycle share at or below the configured stop threshold.');
    rows = RCA_AddKPI(rows, 'Demand Creep Time Share', creepShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Low-speed manoeuvring share between stop and creep thresholds.');
    rows = RCA_AddKPI(rows, 'Demand Urban Time Share', urbanShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand share in the urban-speed band.');
    rows = RCA_AddKPI(rows, 'Demand Medium-Speed Share', mediumSpeedShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand share in the mid-speed or arterial band.');
    rows = RCA_AddKPI(rows, 'Demand Highway Time Share', highwayShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand share above the configured highway-speed threshold.');
    rows = RCA_AddKPI(rows, 'Aggressive Demand Time Share', aggressiveShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Share of the demand cycle with large acceleration or deceleration requests.');
    rows = RCA_AddKPI(rows, 'Desired Acceleration 95th Percentile', positiveAccel95, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Upper-end propulsion demand severity derived from desired speed.');
    rows = RCA_AddKPI(rows, 'Desired Deceleration 95th Percentile', negativeAccel95, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Upper-end braking demand severity derived from desired speed.');
    rows = RCA_AddKPI(rows, 'Planned Stop Count', plannedStopCount, 'count', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Count of demanded stop episodes longer than the configured minimum stop duration.');
    if isfinite(routeDistance_km) && routeDistance_km > 0
        rows = RCA_AddKPI(rows, 'Planned Stop Density', plannedStopCount / routeDistance_km, 'stops/km', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
            'Demand stop count normalized by route distance. ' + distanceBasisNote);
    end

    summary(end + 1) = sprintf(['Demand-cycle severity: average desired speed is %.1f km/h with %.1f%% stop time, %.1f%% urban operation, ', ...
        '%.1f%% highway operation, and %d planned stops.'], mean(desiredSpeed, 'omitnan'), stopShare, urbanShare, highwayShare, plannedStopCount);
end

if roadSlopeSignal.Available && any(isfinite(roadSlope))
    validGradeWeight = isfinite(roadSlope) & isfinite(distanceStep_km) & distanceStep_km >= 0;
    uphillShare = localWeightedShare(roadSlope > config.Thresholds.UphillSlope_pct, distanceStep_km, validGradeWeight);
    downhillShare = localWeightedShare(roadSlope < config.Thresholds.DownhillSlope_pct, distanceStep_km, validGradeWeight);
    steepUphillShare = localWeightedShare(roadSlope > config.Thresholds.SteepSlope_pct, distanceStep_km, validGradeWeight);
    steepDownhillShare = localWeightedShare(roadSlope < -config.Thresholds.SteepSlope_pct, distanceStep_km, validGradeWeight);

    elevationDelta_m = max(distanceStep_km, 0) * 1000 .* roadSlope / 100;
    elevationDelta_m(~isfinite(elevationDelta_m)) = 0;
    elevationGain_m = sum(max(elevationDelta_m, 0), 'omitnan');
    elevationLoss_m = sum(max(-elevationDelta_m, 0), 'omitnan');

    speedForGrade = speedForContext;
    validHighDemand = validGradeWeight & isfinite(speedForGrade);
    highSpeedUphillShare = localWeightedShare(roadSlope > config.Thresholds.UphillSlope_pct & speedForGrade >= config.Thresholds.HighwaySpeed_kmh, distanceStep_km, validHighDemand);
    downhillOpportunityShare = localWeightedShare(roadSlope < config.Thresholds.DownhillSlope_pct & speedForGrade > config.Thresholds.StopSpeed_kmh, distanceStep_km, validHighDemand);

    rows = RCA_AddKPI(rows, 'Mean Road Slope', mean(roadSlope, 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Signed average route slope.');
    rows = RCA_AddKPI(rows, 'Mean Absolute Road Slope', mean(abs(roadSlope), 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Average absolute slope as a grade-severity indicator.');
    rows = RCA_AddKPI(rows, 'Maximum Uphill Slope', max(roadSlope, [], 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Peak uphill grade.');
    rows = RCA_AddKPI(rows, 'Maximum Downhill Slope', min(roadSlope, [], 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Peak downhill grade.');
    rows = RCA_AddKPI(rows, 'Uphill Distance Share', uphillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share with positive grade above the uphill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Downhill Distance Share', downhillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share with negative grade below the downhill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Steep Uphill Distance Share', steepUphillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share above the severe uphill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Steep Downhill Distance Share', steepDownhillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share below the severe downhill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Approximate Elevation Gain', elevationGain_m, 'm', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Approximate cumulative climb from slope and route distance. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Approximate Elevation Loss', elevationLoss_m, 'm', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Approximate cumulative descent from slope and route distance. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'High-Speed Uphill Distance Share', highSpeedUphillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + speedBasisLabel, ...
        'Distance share where route grade and speed demand combine into a high power requirement.');
    rows = RCA_AddKPI(rows, 'Downhill Opportunity Distance Share', downhillOpportunityShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + speedBasisLabel, ...
        'Distance share with downhill motion where recuperation opportunity should exist.');

    summary(end + 1) = sprintf(['Route severity: uphill share is %.1f%%, steep uphill share is %.1f%%, and approximate elevation gain is %.0f m. ', ...
        'High-speed uphill share is %.1f%%.'], uphillShare, steepUphillShare, elevationGain_m, highSpeedUphillShare);
end

if ambientTempSignal.Available && any(isfinite(ambientTemp))
    validAmbient = isfinite(ambientTemp);
    hotShare = 100 * RCA_FractionTrue(ambientTemp >= config.Thresholds.HotAmbient_C, validAmbient);
    coldShare = 100 * RCA_FractionTrue(ambientTemp <= config.Thresholds.ColdAmbient_C, validAmbient);
    mildShare = 100 * RCA_FractionTrue(ambientTemp > config.Thresholds.ColdAmbient_C & ambientTemp < config.Thresholds.HotAmbient_C, validAmbient);

    rows = RCA_AddKPI(rows, 'Mean Ambient Temperature', mean(ambientTemp, 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Trip-average ambient temperature.');
    rows = RCA_AddKPI(rows, 'Minimum Ambient Temperature', min(ambientTemp, [], 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Minimum ambient temperature.');
    rows = RCA_AddKPI(rows, 'Maximum Ambient Temperature', max(ambientTemp, [], 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Maximum ambient temperature.');
    rows = RCA_AddKPI(rows, 'Ambient Temperature Range', max(ambientTemp, [], 'omitnan') - min(ambientTemp, [], 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Ambient temperature spread across the route.');
    rows = RCA_AddKPI(rows, 'Hot Ambient Time Share', hotShare, '%', 'ThermalContext', 'Environment', 'amb_temp', 'Time share above the configured hot ambient threshold.');
    rows = RCA_AddKPI(rows, 'Cold Ambient Time Share', coldShare, '%', 'ThermalContext', 'Environment', 'amb_temp', 'Time share below the configured cold ambient threshold.');
    rows = RCA_AddKPI(rows, 'Mild Ambient Time Share', mildShare, '%', 'ThermalContext', 'Environment', 'amb_temp', 'Time share inside the mild ambient band.');

    summary(end + 1) = sprintf('Thermal context: mean ambient temperature is %.1f degC, with %.1f%% hot exposure and %.1f%% cold exposure.', ...
        mean(ambientTemp, 'omitnan'), hotShare, coldShare);
end

classificationLine = localEnvironmentClassification(summary, desiredSpeedSignal.Available && any(isfinite(desiredSpeed)), ...
    roadSlopeSignal.Available && any(isfinite(roadSlope)), ambientTempSignal.Available && any(isfinite(ambientTemp)), ...
    rows, config);
if strlength(classificationLine) > 0
    summary(end + 1) = classificationLine;
end

if roadSlopeSignal.Available && any(roadSlope > config.Thresholds.SteepSlope_pct)
    recs(end + 1) = "Treat this simulation as route-severe when comparing efficiency or range; slope alone can dominate trip energy and performance demand.";
    evidence(end + 1) = "Steep uphill exposure is present in the route profile.";
end

if desiredSpeedSignal.Available && any(isfinite(desiredSpeed))
    highwayShareValue = localKPIValue(rows, "Demand Highway Time Share");
    stopShareValue = localKPIValue(rows, "Demand Stop Time Share");
    aggressiveShareValue = localKPIValue(rows, "Aggressive Demand Time Share");
    if isfinite(highwayShareValue) && highwayShareValue > 25
        recs(end + 1) = "High-speed demand is material, so aerodynamic load and sustained power capability should be treated as first-order drivers in stakeholder reviews.";
        evidence(end + 1) = sprintf('Highway demand share is %.1f%%.', highwayShareValue);
    end
    if isfinite(stopShareValue) && isfinite(aggressiveShareValue) && stopShareValue > 15 && aggressiveShareValue > 10
        recs(end + 1) = "The duty cycle is stop-go and transient-heavy; regen effectiveness, brake blending, and transient controller behaviour deserve explicit review.";
        evidence(end + 1) = sprintf('Stop share is %.1f%% and aggressive demand share is %.1f%%.', stopShareValue, aggressiveShareValue);
    end
end

if ambientTempSignal.Available && any(isfinite(ambientTemp))
    hotShareValue = localKPIValue(rows, "Hot Ambient Time Share");
    coldShareValue = localKPIValue(rows, "Cold Ambient Time Share");
    if isfinite(hotShareValue) && hotShareValue > 10
        recs(end + 1) = "Correlate hot ambient periods with auxiliary power, battery temperature, and power limiting before assigning efficiency loss only to the propulsion system.";
        evidence(end + 1) = sprintf('Hot ambient exposure is %.1f%% of the trip.', hotShareValue);
    end
    if isfinite(coldShareValue) && coldShareValue > 10
        recs(end + 1) = "Cold-weather exposure is material; inspect warm-up demand, charge acceptance, and auxiliary thermal load as separate contributors.";
        evidence(end + 1) = sprintf('Cold ambient exposure is %.1f%% of the trip.', coldShareValue);
    end
end

if any([desiredSpeedSignal.Available, roadSlopeSignal.Available, ambientTempSignal.Available])
    actualSpeedAvailable = any(isfinite(actualSpeed));
    desiredAccel = localDifferentiateTrace(t, desiredSpeed / 3.6);
    cumulativeElevation_m = cumsum(localFiniteOrZero(max(distanceStep_km, 0) * 1000 .* roadSlope / 100));

    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(4, 1, 1);
    legendEntries = {};
    if desiredSpeedSignal.Available && any(isfinite(desiredSpeed))
        plot(t, desiredSpeed, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
        legendEntries{end + 1} = 'Desired speed';
    end
    if actualSpeedAvailable
        plot(t, actualSpeed, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
        legendEntries{end + 1} = 'Vehicle speed';
    end
    title('Environment Overview: Drive-Cycle Demand and Vehicle Response');
    ylabel('Speed (km/h)');
    if ~isempty(legendEntries)
        legend(legendEntries, 'Location', 'best');
    end
    grid on;

    subplot(4, 1, 2);
    if desiredSpeedSignal.Available && any(isfinite(desiredAccel))
        plot(t, desiredAccel, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
        yline(config.Thresholds.SignificantAccel_mps2, '--', 'Color', config.Plot.Colors.Warning);
        yline(-config.Thresholds.SignificantAccel_mps2, '--', 'Color', config.Plot.Colors.Warning);
        ylabel('Demand acc (m/s^2)');
    else
        text(0.1, 0.5, 'Desired speed unavailable for demand acceleration.', 'Units', 'normalized');
    end
    title('Desired Speed Acceleration Severity');
    grid on;

    subplot(4, 1, 3);
    if roadSlopeSignal.Available && any(isfinite(roadSlope))
        plot(t, roadSlope, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth); hold on;
        yline(config.Thresholds.UphillSlope_pct, '--', 'Color', config.Plot.Colors.Warning);
        yline(config.Thresholds.SteepSlope_pct, ':', 'Color', config.Plot.Colors.Warning);
        yline(config.Thresholds.DownhillSlope_pct, '--', 'Color', config.Plot.Colors.Warning);
        ylabel('Slope (%)');
    else
        text(0.1, 0.5, 'Road slope unavailable.', 'Units', 'normalized');
    end
    title('Route Grade Severity');
    grid on;

    subplot(4, 1, 4);
    if ambientTempSignal.Available && any(isfinite(ambientTemp))
        plot(t, ambientTemp, 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth); hold on;
        yline(config.Thresholds.HotAmbient_C, '--', 'Color', config.Plot.Colors.Warning);
        yline(config.Thresholds.ColdAmbient_C, '--', 'Color', config.Plot.Colors.Warning);
        ylabel('Ambient (degC)');
    else
        text(0.1, 0.5, 'Ambient temperature unavailable.', 'Units', 'normalized');
    end
    xlabel('Time (s)');
    title('Thermal Context');
    grid on;

    plotFiles(end + 1) = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Environment'), 'Environment_Overview', config));
    close(fig);

    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(2, 2, 1);
    if desiredSpeedSignal.Available && any(isfinite(desiredSpeed))
        demandShares = [ ...
            localKPIValue(rows, "Demand Stop Time Share"), ...
            localKPIValue(rows, "Demand Creep Time Share"), ...
            localKPIValue(rows, "Demand Urban Time Share"), ...
            localKPIValue(rows, "Demand Medium-Speed Share"), ...
            localKPIValue(rows, "Demand Highway Time Share")];
        bar(categorical({'Stop', 'Creep', 'Urban', 'Medium', 'Highway'}), demandShares, 'FaceColor', config.Plot.Colors.Demand);
        ylabel('Time share (%)');
        localAddCriteriaBox(gca, sprintf(['Stop: <= %.1f km/h\nCreep: %.1f to %.1f km/h\nUrban: > %.1f to < %.1f km/h\n', ...
            'Medium: >= %.1f to < %.1f km/h\nHighway: >= %.1f km/h'], ...
            config.Thresholds.StopSpeed_kmh, config.Thresholds.StopSpeed_kmh, config.Thresholds.CreepSpeed_kmh, ...
            config.Thresholds.CreepSpeed_kmh, config.Thresholds.UrbanSpeedUpper_kmh, ...
            config.Thresholds.UrbanSpeedUpper_kmh, config.Thresholds.HighwaySpeed_kmh, config.Thresholds.HighwaySpeed_kmh), config);
    else
        text(0.1, 0.5, 'Desired speed unavailable.', 'Units', 'normalized');
    end
    title('Demand-Speed Distribution');
    grid on;

    subplot(2, 2, 2);
    if roadSlopeSignal.Available && any(isfinite(roadSlope))
        gradeShares = [ ...
            localKPIValue(rows, "Uphill Distance Share"), ...
            localKPIValue(rows, "Downhill Distance Share"), ...
            localKPIValue(rows, "Steep Uphill Distance Share"), ...
            localKPIValue(rows, "Steep Downhill Distance Share"), ...
            localKPIValue(rows, "High-Speed Uphill Distance Share")];
        bar(categorical({'Uphill', 'Downhill', 'Steep up', 'Steep down', 'High-speed up'}), gradeShares, 'FaceColor', config.Plot.Colors.Slope);
        ylabel('Distance share (%)');
        localAddCriteriaBox(gca, sprintf(['Uphill: > %.1f %%\nDownhill: < %.1f %%\nSteep up: > %.1f %%\n', ...
            'Steep down: < %.1f %%\nHigh-speed up: uphill and speed >= %.1f km/h'], ...
            config.Thresholds.UphillSlope_pct, config.Thresholds.DownhillSlope_pct, ...
            config.Thresholds.SteepSlope_pct, -config.Thresholds.SteepSlope_pct, config.Thresholds.HighwaySpeed_kmh), config);
    else
        text(0.1, 0.5, 'Road slope unavailable.', 'Units', 'normalized');
    end
    title('Grade Exposure Distribution');
    grid on;

    subplot(2, 2, 3);
    if roadSlopeSignal.Available && any(isfinite(cumulativeElevation_m))
        if any(isfinite(cumulativeDistance_km)) && max(cumulativeDistance_km, [], 'omitnan') > 0
            profileX = cumulativeDistance_km;
            xlabel('Distance (km)');
        else
            profileX = t;
            xlabel('Time (s)');
        end

        if ambientTempSignal.Available && any(isfinite(ambientTemp))
            validProfile = isfinite(profileX) & isfinite(cumulativeElevation_m) & isfinite(ambientTemp);
            plot(profileX, cumulativeElevation_m, 'Color', config.Plot.Colors.Neutral, 'LineWidth', max(config.Plot.LineWidth - 0.2, 1.0)); hold on;
            scatter(profileX(validProfile), cumulativeElevation_m(validProfile), 16, ambientTemp(validProfile), 'filled');
            cb = colorbar;
            cb.Label.String = 'Ambient temperature (degC)';
        else
            plot(profileX, cumulativeElevation_m, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
        end
        ylabel('Cumulative elevation (m)');
    else
        text(0.1, 0.5, 'Insufficient route grade data.', 'Units', 'normalized');
    end
    title('Approximate Route Elevation Profile');
    grid on;

    subplot(2, 2, 4);
    if roadSlopeSignal.Available && any(isfinite(roadSlope)) && any(isfinite(speedForContext))
        validScatter = isfinite(speedForContext) & isfinite(roadSlope);
        scatter(speedForContext(validScatter), roadSlope(validScatter), 14, 'filled', ...
            'MarkerFaceColor', config.Plot.Colors.Demand, 'MarkerEdgeColor', config.Plot.Colors.Neutral);
        xlabel('Demand speed (km/h)');
        ylabel('Slope (%)');
    else
        text(0.1, 0.5, 'Need speed and slope for combined severity map.', 'Units', 'normalized');
    end
    title('Speed-Demand Versus Grade Severity');
    grid on;

    plotFiles(end + 1) = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Environment'), 'Environment_Severity_Map', config));
    close(fig);
end

result.Available = ~isempty(rows);
result.KPITable = localRemoveSubsystemColumn(RCA_FinalizeKPITable(rows));
result.FigureFiles = plotFiles;
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Environment", recs, evidence);
end

function [distanceStep_km, cumulativeDistance_km, distanceBasisLabel, distanceBasisNote] = localDistanceBasis(derived, speedBasis_kmh, dt_s, config)
distanceStep_km = derived.distanceStep_km(:);
if isempty(distanceStep_km)
    distanceStep_km = NaN(size(speedBasis_kmh(:)));
end

distanceStep_km = double(distanceStep_km(:));
speedBasis_kmh = double(speedBasis_kmh(:));
dt_s = double(dt_s(:));

actualDistance_km = sum(max(distanceStep_km, 0), 'omitnan');
if ~isfinite(actualDistance_km) || actualDistance_km < config.General.MinimumDistanceForWhpkm_km
    demandDistanceStep_km = max(speedBasis_kmh, 0) .* max(dt_s, 0) / 3600;
    if any(isfinite(demandDistanceStep_km)) && sum(demandDistanceStep_km, 'omitnan') > 0
        distanceStep_km = demandDistanceStep_km;
        distanceBasisLabel = "veh_des_vel";
        distanceBasisNote = "Route distance was approximated by integrating desired speed because actual vehicle distance was unavailable or too small.";
    else
        distanceBasisLabel = "veh_pos or veh_vel";
        distanceBasisNote = "Route distance basis was unavailable, so distance-weighted KPI confidence is limited.";
    end
else
    distanceBasisLabel = "veh_pos or veh_vel";
    distanceBasisNote = "Route distance uses vehicle position or vehicle speed derived distance.";
end

distanceStep_km(~isfinite(distanceStep_km)) = 0;
cumulativeDistance_km = cumsum(max(distanceStep_km, 0));
end

function deriv = localDifferentiateTrace(t, x)
deriv = NaN(size(x));
t = double(t(:));
x = double(x(:));
commonLength = min(numel(t), numel(x));
t = t(1:commonLength);
x = x(1:commonLength);
deriv = NaN(commonLength, 1);

for iPoint = 2:commonLength
    if isfinite(t(iPoint)) && isfinite(t(iPoint - 1)) && isfinite(x(iPoint)) && isfinite(x(iPoint - 1))
        dt = t(iPoint) - t(iPoint - 1);
        if dt > 0
            deriv(iPoint) = (x(iPoint) - x(iPoint - 1)) / dt;
        end
    end
end
end

function count = localCountEpisodes(mask, t, minDuration_s)
mask = logical(mask(:));
t = double(t(:));
commonLength = min(numel(mask), numel(t));
mask = mask(1:commonLength);
t = t(1:commonLength);
mask(~isfinite(t)) = false;

transitions = diff([false; mask; false]);
startIdx = find(transitions == 1);
endIdx = find(transitions == -1) - 1;
count = 0;

for iEpisode = 1:numel(startIdx)
    if endIdx(iEpisode) > startIdx(iEpisode)
        duration_s = t(endIdx(iEpisode)) - t(startIdx(iEpisode));
    else
        duration_s = 0;
    end
    if isfinite(duration_s) && duration_s >= minDuration_s
        count = count + 1;
    end
end
end

function share = localWeightedShare(mask, weights, validMask)
mask = logical(mask(:));
weights = double(weights(:));
validMask = logical(validMask(:));
commonLength = min([numel(mask), numel(weights), numel(validMask)]);
mask = mask(1:commonLength);
weights = weights(1:commonLength);
validMask = validMask(1:commonLength);

weights(~isfinite(weights)) = 0;
weights(weights < 0) = 0;
validMask = validMask & isfinite(weights);
totalWeight = sum(weights(validMask));

if totalWeight <= 0
    share = NaN;
else
    share = 100 * sum(weights(mask & validMask)) / totalWeight;
end
end

function values = localFiniteOrZero(values)
values = double(values);
values(~isfinite(values)) = 0;
end

function value = localKPIValue(rows, kpiName)
value = NaN;
if isempty(rows)
    return;
end

for iRow = 1:size(rows, 1)
    if strcmp(rows{iRow, 1}, kpiName)
        value = rows{iRow, 2};
        return;
    end
end
end

function line = localEnvironmentClassification(summary, hasDemand, hasGrade, hasAmbient, rows, config)
line = "";
if ~(hasDemand || hasGrade || hasAmbient)
    return;
end

mobilityClass = "mixed-speed";
gradeClass = "moderate-grade";
thermalClass = "mild-temperature";

if hasDemand
    stopShare = localKPIValue(rows, "Demand Stop Time Share");
    highwayShare = localKPIValue(rows, "Demand Highway Time Share");
    if isfinite(stopShare) && stopShare > 25
        mobilityClass = "stop-go urban";
    elseif isfinite(highwayShare) && highwayShare > 35
        mobilityClass = "high-speed";
    end
end

if hasGrade
    steepUphillShare = localKPIValue(rows, "Steep Uphill Distance Share");
    uphillShare = localKPIValue(rows, "Uphill Distance Share");
    if isfinite(steepUphillShare) && steepUphillShare > 10
        gradeClass = "hilly";
    elseif isfinite(uphillShare) && uphillShare < 10
        gradeClass = "mostly flat";
    end
end

if hasAmbient
    hotShare = localKPIValue(rows, "Hot Ambient Time Share");
    coldShare = localKPIValue(rows, "Cold Ambient Time Share");
    if isfinite(hotShare) && hotShare > 10
        thermalClass = "hot-weather";
    elseif isfinite(coldShare) && coldShare > 10
        thermalClass = "cold-weather";
    end
end

line = sprintf(['Environment classification: %s, %s, %s route. This context should be used before comparing ', ...
    'efficiency, range, or performance across simulations. Thresholds are editable heuristics from RCA_Config.'], ...
    mobilityClass, gradeClass, thermalClass);
end

function localAddCriteriaBox(axisHandle, noteText, config)
text(axisHandle, 0.98, 0.98, noteText, ...
    'Units', 'normalized', ...
    'HorizontalAlignment', 'right', ...
    'VerticalAlignment', 'top', ...
    'FontSize', max(config.Plot.FontSize - 2, 8), ...
    'BackgroundColor', [1 1 1], ...
    'EdgeColor', config.Plot.Colors.Neutral, ...
    'Margin', 4);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', localRemoveSubsystemColumn(RCA_FinalizeKPITable([])), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end

function kpiTable = localRemoveSubsystemColumn(kpiTable)
if istable(kpiTable) && ismember('Subsystem', kpiTable.Properties.VariableNames)
    kpiTable.Subsystem = [];
end
end
