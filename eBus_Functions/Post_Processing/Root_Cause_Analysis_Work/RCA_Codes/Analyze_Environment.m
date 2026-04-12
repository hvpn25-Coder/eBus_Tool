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

desiredAccel = localDifferentiateTrace(t, desiredSpeed / 3.6);

if desiredSpeedSignal.Available && any(isfinite(desiredSpeed))
    validDesired = isfinite(desiredSpeed);
    validDesiredAccel = isfinite(desiredAccel);

    stopShare = 100 * RCA_FractionTrue(desiredSpeed <= config.Thresholds.StopSpeed_kmh, validDesired);
    creepShare = 100 * RCA_FractionTrue(desiredSpeed > config.Thresholds.StopSpeed_kmh & desiredSpeed <= config.Thresholds.CreepSpeed_kmh, validDesired);
    urbanShare = 100 * RCA_FractionTrue(desiredSpeed > config.Thresholds.CreepSpeed_kmh & desiredSpeed < config.Thresholds.UrbanSpeedUpper_kmh, validDesired);
    mediumSpeedShare = 100 * RCA_FractionTrue(desiredSpeed >= config.Thresholds.UrbanSpeedUpper_kmh & desiredSpeed < config.Thresholds.HighwaySpeed_kmh, validDesired);
    highwayShare = 100 * RCA_FractionTrue(desiredSpeed >= config.Thresholds.HighwaySpeed_kmh, validDesired);
    aggressiveShare = 100 * RCA_FractionTrue(abs(desiredAccel) >= config.Thresholds.SignificantAccel_mps2, validDesiredAccel);
    plannedStopCount = localCountEpisodes(desiredSpeed <= config.Thresholds.StopSpeed_kmh, t, config.Thresholds.MinStopDuration_s);
    movingSpeed = desiredSpeed(desiredSpeed > config.Thresholds.StopSpeed_kmh & isfinite(desiredSpeed));
    positiveAccel95 = RCA_Percentile(desiredAccel(desiredAccel > 0), 95);
    negativeAccel95 = RCA_Percentile(-desiredAccel(desiredAccel < 0), 95);
    maxDesiredAccel = max(desiredAccel, [], 'omitnan');
    maxDesiredDecel = min(desiredAccel, [], 'omitnan');
    meanPositiveAccel = mean(desiredAccel(desiredAccel > 0), 'omitnan');
    meanNegativeDecel = mean(-desiredAccel(desiredAccel < 0), 'omitnan');
    startStopCycleCount = localCountRisingEdges(desiredSpeed > config.Thresholds.StopSpeed_kmh, t);
    demandIndex = localDemandSeverityIndex(stopShare, highwayShare, aggressiveShare, plannedStopCount, routeDistance_km);

    rows = RCA_AddKPI(rows, 'Average Desired Speed', mean(desiredSpeed, 'omitnan'), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', 'Average demanded vehicle speed from the environment subsystem.');
    rows = RCA_AddKPI(rows, 'Average Moving Desired Speed', mean(movingSpeed, 'omitnan'), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', ...
        'Average desired speed while the demand cycle is above the stop threshold. This separates route pace from dwell/stop exposure.');
    rows = RCA_AddKPI(rows, 'Peak Desired Speed', max(desiredSpeed, [], 'omitnan'), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', 'Peak demanded speed.');
    rows = RCA_AddKPI(rows, 'Desired Speed 95th Percentile', RCA_Percentile(desiredSpeed, 95), 'km/h', 'DemandCycle', 'Environment', 'veh_des_vel', 'Upper-end speed demand severity.');
    rows = RCA_AddKPI(rows, 'Demand Stop Time Share', stopShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand-cycle share at or below the configured stop threshold.');
    rows = RCA_AddKPI(rows, 'Demand Creep Time Share', creepShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Low-speed manoeuvring share between stop and creep thresholds.');
    rows = RCA_AddKPI(rows, 'Demand Urban Time Share', urbanShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand share in the urban-speed band.');
    rows = RCA_AddKPI(rows, 'Demand Medium-Speed Share', mediumSpeedShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand share in the mid-speed or arterial band.');
    rows = RCA_AddKPI(rows, 'Demand Highway Time Share', highwayShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel', 'Demand share above the configured highway-speed threshold.');
    rows = RCA_AddKPI(rows, 'Aggressive Demand Time Share', aggressiveShare, '%', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Share of the demand cycle with large acceleration or deceleration requests.');
    rows = RCA_AddKPI(rows, 'Maximum Desired Acceleration', maxDesiredAccel, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
        'Maximum positive desired acceleration. This indicates peak propulsion demand imposed by the route schedule.');
    rows = RCA_AddKPI(rows, 'Maximum Desired Deceleration', maxDesiredDecel, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
        'Maximum negative desired acceleration. This indicates peak braking or regen opportunity imposed by the route schedule.');
    rows = RCA_AddKPI(rows, 'Mean Positive Desired Acceleration', meanPositiveAccel, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
        'Average positive desired acceleration during acceleration portions of the route demand.');
    rows = RCA_AddKPI(rows, 'Mean Desired Deceleration Magnitude', meanNegativeDecel, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
        'Average deceleration magnitude during braking portions of the route demand.');
    rows = RCA_AddKPI(rows, 'Desired Acceleration 95th Percentile', positiveAccel95, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Upper-end propulsion demand severity derived from desired speed.');
    rows = RCA_AddKPI(rows, 'Desired Deceleration 95th Percentile', negativeAccel95, 'm/s^2', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Upper-end braking demand severity derived from desired speed.');
    rows = RCA_AddKPI(rows, 'Planned Stop Count', plannedStopCount, 'count', 'DemandCycle', 'Environment', 'veh_des_vel + time', 'Count of demanded stop episodes longer than the configured minimum stop duration.');
    rows = RCA_AddKPI(rows, 'Start-Stop Cycle Count', startStopCycleCount, 'count', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
        'Count of transitions from stop to moving demand. This is a stakeholder-friendly stop-go intensity indicator.');
    if isfinite(routeDistance_km) && routeDistance_km > 0
        rows = RCA_AddKPI(rows, 'Planned Stop Density', plannedStopCount / routeDistance_km, 'stops/km', 'DemandCycle', 'Environment', 'veh_des_vel + time', ...
            'Demand stop count normalized by route distance. ' + distanceBasisNote);
    end
    rows = RCA_AddKPI(rows, 'Route Demand Severity Index', demandIndex, '0-100', 'CombinedEnvironmentSeverity', 'Environment', 'veh_des_vel + time', ...
        'Heuristic index combining stop exposure, high-speed exposure, transient acceleration share, and stop density. Higher means a more demanding duty cycle.');

    summary(end + 1) = sprintf(['Demand-cycle severity: average desired speed is %.1f km/h with %.1f%% stop time, %.1f%% urban operation, ', ...
        '%.1f%% highway operation, %d planned stops, and a %.0f/100 route demand severity index.'], ...
        mean(desiredSpeed, 'omitnan'), stopShare, urbanShare, highwayShare, plannedStopCount, demandIndex);
end

if roadSlopeSignal.Available && any(isfinite(roadSlope))
    validGradeWeight = isfinite(roadSlope) & isfinite(distanceStep_km) & distanceStep_km >= 0;
    validGradeTime = isfinite(roadSlope) & isfinite(t);
    uphillShare = localWeightedShare(roadSlope > config.Thresholds.UphillSlope_pct, distanceStep_km, validGradeWeight);
    downhillShare = localWeightedShare(roadSlope < config.Thresholds.DownhillSlope_pct, distanceStep_km, validGradeWeight);
    steepUphillShare = localWeightedShare(roadSlope > config.Thresholds.SteepSlope_pct, distanceStep_km, validGradeWeight);
    steepDownhillShare = localWeightedShare(roadSlope < -config.Thresholds.SteepSlope_pct, distanceStep_km, validGradeWeight);
    uphillTimeShare = localTimeShare(roadSlope > config.Thresholds.UphillSlope_pct, t, validGradeTime);
    downhillTimeShare = localTimeShare(roadSlope < config.Thresholds.DownhillSlope_pct, t, validGradeTime);
    moderateGradeDistanceShare = localWeightedShare(abs(roadSlope) >= config.Thresholds.EnvironmentModerateSlope_pct, distanceStep_km, validGradeWeight);
    severeGradeDistanceShare = localWeightedShare(abs(roadSlope) >= config.Thresholds.EnvironmentSevereSlope_pct, distanceStep_km, validGradeWeight);

    elevationDelta_m = max(distanceStep_km, 0) * 1000 .* roadSlope / 100;
    elevationDelta_m(~isfinite(elevationDelta_m)) = 0;
    elevationGain_m = sum(max(elevationDelta_m, 0), 'omitnan');
    elevationLoss_m = sum(max(-elevationDelta_m, 0), 'omitnan');
    netElevation_m = sum(elevationDelta_m, 'omitnan');
    hillinessIndex_mPerKm = localSafeDivide(elevationGain_m + elevationLoss_m, routeDistance_km);

    speedForGrade = speedForContext;
    validHighDemand = validGradeWeight & isfinite(speedForGrade);
    highSpeedUphillShare = localWeightedShare(roadSlope > config.Thresholds.UphillSlope_pct & speedForGrade >= config.Thresholds.HighwaySpeed_kmh, distanceStep_km, validHighDemand);
    downhillOpportunityShare = localWeightedShare(roadSlope < config.Thresholds.DownhillSlope_pct & speedForGrade > config.Thresholds.StopSpeed_kmh, distanceStep_km, validHighDemand);
    gradeSeverityIndex = localGradeSeverityIndex(mean(abs(roadSlope), 'omitnan'), uphillShare, steepUphillShare, hillinessIndex_mPerKm, config);
    regenOpportunityIndex = localRegenOpportunityIndex(downhillShare, steepDownhillShare, elevationLoss_m, routeDistance_km);

    rows = RCA_AddKPI(rows, 'Mean Road Slope', mean(roadSlope, 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Signed average route slope.');
    rows = RCA_AddKPI(rows, 'Mean Absolute Road Slope', mean(abs(roadSlope), 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Average absolute slope as a grade-severity indicator.');
    rows = RCA_AddKPI(rows, 'Maximum Uphill Slope', max(roadSlope, [], 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Peak uphill grade.');
    rows = RCA_AddKPI(rows, 'Maximum Downhill Slope', min(roadSlope, [], 'omitnan'), '%', 'RouteSeverity', 'Environment', 'road_slp', 'Peak downhill grade.');
    rows = RCA_AddKPI(rows, 'Uphill Time Share', uphillTimeShare, '%', 'RouteSeverity', 'Environment', 'road_slp + time', ...
        'Time share above the uphill threshold. This explains how long the bus is exposed to grade load independent of distance basis.');
    rows = RCA_AddKPI(rows, 'Downhill Time Share', downhillTimeShare, '%', 'RouteSeverity', 'Environment', 'road_slp + time', ...
        'Time share below the downhill threshold. This indicates how long regenerative opportunity may be available.');
    rows = RCA_AddKPI(rows, 'Uphill Distance Share', uphillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share with positive grade above the uphill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Downhill Distance Share', downhillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share with negative grade below the downhill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Steep Uphill Distance Share', steepUphillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share above the severe uphill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Steep Downhill Distance Share', steepDownhillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share below the severe downhill threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Moderate Grade Distance Share', moderateGradeDistanceShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share with absolute grade above the moderate stakeholder threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Severe Grade Distance Share', severeGradeDistanceShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Distance share with absolute grade above the severe stakeholder threshold. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Approximate Elevation Gain', elevationGain_m, 'm', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Approximate cumulative climb from slope and route distance. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Approximate Elevation Loss', elevationLoss_m, 'm', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Approximate cumulative descent from slope and route distance. ' + distanceBasisNote);
    rows = RCA_AddKPI(rows, 'Approximate Net Elevation Change', netElevation_m, 'm', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Approximate end-to-end elevation tendency from slope and route distance. Positive indicates net climb; negative indicates net descent.');
    rows = RCA_AddKPI(rows, 'Hilliness Index', hillinessIndex_mPerKm, 'm/km', 'RouteSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Cumulative climb plus descent normalized by route distance. This is a compact topography-severity KPI.');
    rows = RCA_AddKPI(rows, 'High-Speed Uphill Distance Share', highSpeedUphillShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + speedBasisLabel, ...
        'Distance share where route grade and speed demand combine into a high power requirement.');
    rows = RCA_AddKPI(rows, 'Downhill Opportunity Distance Share', downhillOpportunityShare, '%', 'RouteSeverity', 'Environment', "road_slp + " + speedBasisLabel, ...
        'Distance share with downhill motion where recuperation opportunity should exist.');
    rows = RCA_AddKPI(rows, 'Grade Severity Index', gradeSeverityIndex, '0-100', 'CombinedEnvironmentSeverity', 'Environment', "road_slp + " + distanceBasisLabel, ...
        'Heuristic grade severity score combining mean absolute slope, uphill share, steep uphill exposure, and hilliness index.');
    rows = RCA_AddKPI(rows, 'Regen Opportunity Index', regenOpportunityIndex, '0-100', 'CombinedEnvironmentSeverity', 'Environment', "road_slp + " + speedBasisLabel, ...
        'Heuristic opportunity score from downhill distance exposure, steep downhill exposure, and approximate descent per kilometre.');

    summary(end + 1) = sprintf(['Route severity: uphill share is %.1f%%, steep uphill share is %.1f%%, and approximate elevation gain is %.0f m. ', ...
        'High-speed uphill share is %.1f%%, hilliness is %.1f m/km, and grade severity index is %.0f/100.'], ...
        uphillShare, steepUphillShare, elevationGain_m, highSpeedUphillShare, hillinessIndex_mPerKm, gradeSeverityIndex);
end

if ambientTempSignal.Available && any(isfinite(ambientTemp))
    validAmbient = isfinite(ambientTemp);
    hotShare = 100 * RCA_FractionTrue(ambientTemp >= config.Thresholds.HotAmbient_C, validAmbient);
    coldShare = 100 * RCA_FractionTrue(ambientTemp <= config.Thresholds.ColdAmbient_C, validAmbient);
    mildShare = 100 * RCA_FractionTrue(ambientTemp > config.Thresholds.ColdAmbient_C & ambientTemp < config.Thresholds.HotAmbient_C, validAmbient);
    freezingShare = 100 * RCA_FractionTrue(ambientTemp <= config.Thresholds.EnvironmentFreezingAmbient_C, validAmbient);
    coolShare = 100 * RCA_FractionTrue(ambientTemp <= config.Thresholds.EnvironmentCoolAmbient_C, validAmbient);
    comfortShare = 100 * RCA_FractionTrue(ambientTemp >= config.Thresholds.EnvironmentMildAmbientLow_C & ambientTemp <= config.Thresholds.EnvironmentMildAmbientHigh_C, validAmbient);
    warmShare = 100 * RCA_FractionTrue(ambientTemp >= config.Thresholds.EnvironmentWarmAmbient_C, validAmbient);
    extremeHotShare = 100 * RCA_FractionTrue(ambientTemp >= config.Thresholds.EnvironmentExtremeHotAmbient_C, validAmbient);
    ambientRate_CPerMin = localDifferentiateTrace(t, ambientTemp) * 60;
    ambientVariation95 = RCA_Percentile(abs(ambientRate_CPerMin), 95);
    coldExposure_degCh = localExposureIntegral(ambientTemp, t, config.Thresholds.EnvironmentMildAmbientLow_C, "below");
    hotExposure_degCh = localExposureIntegral(ambientTemp, t, config.Thresholds.EnvironmentMildAmbientHigh_C, "above");
    climateSeverityIndex = localClimateSeverityIndex(coldShare, hotShare, comfortShare, coldExposure_degCh, hotExposure_degCh);

    rows = RCA_AddKPI(rows, 'Mean Ambient Temperature', mean(ambientTemp, 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Trip-average ambient temperature.');
    rows = RCA_AddKPI(rows, 'Minimum Ambient Temperature', min(ambientTemp, [], 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Minimum ambient temperature.');
    rows = RCA_AddKPI(rows, 'Maximum Ambient Temperature', max(ambientTemp, [], 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Maximum ambient temperature.');
    rows = RCA_AddKPI(rows, 'Ambient Temperature Range', max(ambientTemp, [], 'omitnan') - min(ambientTemp, [], 'omitnan'), 'degC', 'ThermalContext', 'Environment', 'amb_temp', 'Ambient temperature spread across the route.');
    rows = RCA_AddKPI(rows, 'Hot Ambient Time Share', hotShare, '%', 'ThermalContext', 'Environment', 'amb_temp', 'Time share above the configured hot ambient threshold.');
    rows = RCA_AddKPI(rows, 'Cold Ambient Time Share', coldShare, '%', 'ThermalContext', 'Environment', 'amb_temp', 'Time share below the configured cold ambient threshold.');
    rows = RCA_AddKPI(rows, 'Mild Ambient Time Share', mildShare, '%', 'ThermalContext', 'Environment', 'amb_temp', 'Time share inside the mild ambient band.');
    rows = RCA_AddKPI(rows, 'Freezing Ambient Time Share', freezingShare, '%', 'ThermalContext', 'Environment', 'amb_temp', ...
        'Time share at or below the freezing exposure threshold. This can affect battery performance, heating load, and regen acceptance.');
    rows = RCA_AddKPI(rows, 'Cool Ambient Time Share', coolShare, '%', 'ThermalContext', 'Environment', 'amb_temp', ...
        'Time share below the cool-weather threshold. This supports HVAC and battery warm-up interpretation.');
    rows = RCA_AddKPI(rows, 'Comfort Ambient Time Share', comfortShare, '%', 'ThermalContext', 'Environment', 'amb_temp', ...
        'Time share inside the stakeholder mild/comfort temperature band.');
    rows = RCA_AddKPI(rows, 'Warm Ambient Time Share', warmShare, '%', 'ThermalContext', 'Environment', 'amb_temp', ...
        'Time share above the warm-weather threshold where cooling or HVAC demand may become relevant.');
    rows = RCA_AddKPI(rows, 'Extreme Hot Ambient Time Share', extremeHotShare, '%', 'ThermalContext', 'Environment', 'amb_temp', ...
        'Time share above the extreme-hot threshold where thermal derating risk should be reviewed.');
    rows = RCA_AddKPI(rows, 'Cold Exposure Index', coldExposure_degCh, 'degC*h', 'ThermalContext', 'Environment', 'amb_temp + time', ...
        'Integrated degrees below the mild ambient lower bound. Higher values indicate stronger cold-weather burden.');
    rows = RCA_AddKPI(rows, 'Hot Exposure Index', hotExposure_degCh, 'degC*h', 'ThermalContext', 'Environment', 'amb_temp + time', ...
        'Integrated degrees above the mild ambient upper bound. Higher values indicate stronger cooling or thermal-stress burden.');
    rows = RCA_AddKPI(rows, 'Ambient Variation 95th Percentile', ambientVariation95, 'degC/min', 'ThermalContext', 'Environment', 'amb_temp + time', ...
        'Upper percentile of ambient temperature rate of change. High values can reveal route-zone transitions or logging artefacts.');
    rows = RCA_AddKPI(rows, 'Climate Severity Index', climateSeverityIndex, '0-100', 'CombinedEnvironmentSeverity', 'Environment', 'amb_temp + time', ...
        'Heuristic climate severity score combining hot/cold exposure, comfort-band exposure, and degree-hour burden.');

    summary(end + 1) = sprintf(['Thermal context: mean ambient temperature is %.1f degC, with %.1f%% hot exposure, %.1f%% cold exposure, ', ...
        '%.1f%% comfort-band exposure, and climate severity index %.0f/100.'], ...
        mean(ambientTemp, 'omitnan'), hotShare, coldShare, comfortShare, climateSeverityIndex);
end

classificationLine = localEnvironmentClassification(summary, desiredSpeedSignal.Available && any(isfinite(desiredSpeed)), ...
    roadSlopeSignal.Available && any(isfinite(roadSlope)), ambientTempSignal.Available && any(isfinite(ambientTemp)), ...
    rows, config);
if strlength(classificationLine) > 0
    summary(end + 1) = classificationLine;
end

overallSeverityIndex = localOverallEnvironmentSeverityIndex(rows);
if isfinite(overallSeverityIndex)
    rows = RCA_AddKPI(rows, 'Overall Environment Severity Index', overallSeverityIndex, '0-100', 'CombinedEnvironmentSeverity', 'Environment', ...
        'veh_des_vel + road_slp + amb_temp', ...
        'Heuristic combined severity score from route demand, grade severity, climate severity, and regen opportunity context. Higher means harsher external operating conditions.');
    summary(end + 1) = sprintf(['Overall environment severity index is %.0f/100. Use this as a scenario-context score, not as a vehicle performance score; ', ...
        'it explains whether efficiency, range, thermal stress, or performance results were evaluated under mild or harsh external conditions.'], overallSeverityIndex);
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

if isfinite(overallSeverityIndex) && overallSeverityIndex >= 60
    recs(end + 1) = "Classify this run as externally severe before assigning poor efficiency or performance to a subsystem; compare against a mild-route baseline if available.";
    evidence(end + 1) = sprintf('Overall environment severity index is %.0f/100.', overallSeverityIndex);
end

if roadSlopeSignal.Available && desiredSpeedSignal.Available
    highSpeedUphillValue = localKPIValue(rows, "High-Speed Uphill Distance Share");
    if isfinite(highSpeedUphillValue) && highSpeedUphillValue > 5
        recs(end + 1) = "Review high-speed uphill intervals as peak mission-demand zones; correlate them with battery power, motor torque limits, gearbox behavior, and speed tracking.";
        evidence(end + 1) = sprintf('High-speed uphill distance share is %.1f%%.', highSpeedUphillValue);
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

    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(2, 2, 1);
    indexNames = {};
    indexValues = [];
    candidateIndexNames = ["Route Demand Severity Index", "Grade Severity Index", "Climate Severity Index", "Regen Opportunity Index", "Overall Environment Severity Index"];
    for iIndex = 1:numel(candidateIndexNames)
        candidateValue = localKPIValue(rows, candidateIndexNames(iIndex));
        if isfinite(candidateValue)
            indexNames{end + 1} = char(candidateIndexNames(iIndex)); %#ok<AGROW>
            indexValues(end + 1) = candidateValue; %#ok<AGROW>
        end
    end
    if ~isempty(indexValues)
        bar(categorical(indexNames), indexValues, 'FaceColor', config.Plot.Colors.Warning);
        ylim([0 100]);
        ylabel('Index (0-100)');
        title('Environment Severity Index Summary');
        grid on;
    else
        text(0.1, 0.5, 'No combined environment indices available.', 'Units', 'normalized');
        axis off;
    end

    subplot(2, 2, 2);
    if ambientTempSignal.Available && any(isfinite(ambientTemp))
        ambientShares = [ ...
            localKPIValue(rows, "Freezing Ambient Time Share"), ...
            localKPIValue(rows, "Cool Ambient Time Share"), ...
            localKPIValue(rows, "Comfort Ambient Time Share"), ...
            localKPIValue(rows, "Warm Ambient Time Share"), ...
            localKPIValue(rows, "Extreme Hot Ambient Time Share")];
        bar(categorical({'Freezing', 'Cool', 'Comfort', 'Warm', 'Extreme hot'}), ambientShares, 'FaceColor', config.Plot.Colors.Auxiliary);
        ylabel('Time share (%)');
        title('Ambient Exposure Distribution');
        localAddCriteriaBox(gca, sprintf(['Freezing: <= %.1f degC\nCool: <= %.1f degC\nComfort: %.1f to %.1f degC\n', ...
            'Warm: >= %.1f degC\nExtreme hot: >= %.1f degC'], ...
            config.Thresholds.EnvironmentFreezingAmbient_C, config.Thresholds.EnvironmentCoolAmbient_C, ...
            config.Thresholds.EnvironmentMildAmbientLow_C, config.Thresholds.EnvironmentMildAmbientHigh_C, ...
            config.Thresholds.EnvironmentWarmAmbient_C, config.Thresholds.EnvironmentExtremeHotAmbient_C), config);
        grid on;
    else
        text(0.1, 0.5, 'Ambient temperature unavailable.', 'Units', 'normalized');
        axis off;
    end

    subplot(2, 2, 3);
    hold on;
    plottedCumulative = false;
    cumulativeLegend = {};
    if roadSlopeSignal.Available && exist('elevationDelta_m', 'var')
        plot(t, cumsum(max(elevationDelta_m, 0)), 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
        plottedCumulative = true;
        cumulativeLegend{end + 1} = 'Climb (m)';
    end
    if ambientTempSignal.Available && any(isfinite(ambientTemp))
        coldCumulative = localCumulativeExposure(ambientTemp, t, config.Thresholds.EnvironmentMildAmbientLow_C, "below");
        hotCumulative = localCumulativeExposure(ambientTemp, t, config.Thresholds.EnvironmentMildAmbientHigh_C, "above");
        plot(t, coldCumulative, '--', 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
        plot(t, hotCumulative, ':', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
        plottedCumulative = true;
        cumulativeLegend{end + 1} = 'Cold exposure (degC*h)';
        cumulativeLegend{end + 1} = 'Hot exposure (degC*h)';
    end
    if plottedCumulative
        xlabel('Time (s)');
        ylabel('Cumulative burden');
        title('Cumulative Route / Climate Burden');
        legend(cumulativeLegend, 'Location', 'best');
        grid on;
    else
        text(0.1, 0.5, 'No cumulative severity data available.', 'Units', 'normalized');
        axis off;
    end

    subplot(2, 2, 4);
    [timelineClass, classLabels] = localEnvironmentTimelineClass(desiredSpeed, desiredAccel, roadSlope, ambientTemp, config);
    validTimeline = isfinite(timelineClass) & isfinite(t);
    if any(validTimeline)
        stairs(t(validTimeline), timelineClass(validTimeline), 'LineWidth', config.Plot.LineWidth, 'Color', config.Plot.Colors.Demand);
        yticks(1:numel(classLabels));
        yticklabels(cellstr(classLabels));
        ylim([0.5 numel(classLabels) + 0.5]);
        xlabel('Time (s)');
        title('Environment Mission Context Timeline');
        grid on;
    else
        text(0.1, 0.5, 'Insufficient data for environment context timeline.', 'Units', 'normalized');
        axis off;
    end

    plotFiles(end + 1) = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Environment'), 'Environment_Stakeholder_Dashboard', config));
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

function share = localTimeShare(mask, t, validMask)
mask = logical(mask(:));
t = double(t(:));
validMask = logical(validMask(:));
commonLength = min([numel(mask), numel(t), numel(validMask)]);
mask = mask(1:commonLength);
t = t(1:commonLength);
validMask = validMask(1:commonLength);

dt = localSampleDurations(t);
validMask = validMask & isfinite(dt) & dt >= 0;
totalTime = sum(dt(validMask), 'omitnan');
if totalTime <= 0
    share = NaN;
else
    share = 100 * sum(dt(mask & validMask), 'omitnan') / totalTime;
end
end

function count = localCountRisingEdges(mask, t)
mask = logical(mask(:));
t = double(t(:));
commonLength = min(numel(mask), numel(t));
mask = mask(1:commonLength);
t = t(1:commonLength);
mask(~isfinite(t)) = false;
count = sum(diff([false; mask]) == 1);
end

function dt = localSampleDurations(t)
t = double(t(:));
dt = zeros(size(t));
if numel(t) < 2
    return;
end
rawDt = diff(t);
rawDt(~isfinite(rawDt) | rawDt < 0) = 0;
dt(1:end-1) = rawDt;
dt(end) = median(rawDt(rawDt > 0), 'omitnan');
if ~isfinite(dt(end))
    dt(end) = 0;
end
end

function value = localSafeDivide(numerator, denominator)
if isfinite(numerator) && isfinite(denominator) && abs(denominator) > eps
    value = numerator / denominator;
else
    value = NaN;
end
end

function value = localClip(value, lowerLimit, upperLimit)
value = min(max(value, lowerLimit), upperLimit);
end

function score = localDemandSeverityIndex(stopShare, highwayShare, aggressiveShare, stopCount, routeDistance_km)
stopDensity = localSafeDivide(stopCount, routeDistance_km);
components = [ ...
    localClip(stopShare / 40, 0, 1), ...
    localClip(highwayShare / 50, 0, 1), ...
    localClip(aggressiveShare / 30, 0, 1), ...
    localClip(stopDensity / 3, 0, 1)];
score = 100 * mean(components(isfinite(components)), 'omitnan');
end

function score = localGradeSeverityIndex(meanAbsSlope, uphillShare, steepUphillShare, hillinessIndex_mPerKm, config)
components = [ ...
    localClip(meanAbsSlope / max(config.Thresholds.EnvironmentModerateSlope_pct, eps), 0, 1), ...
    localClip(uphillShare / 60, 0, 1), ...
    localClip(steepUphillShare / 25, 0, 1), ...
    localClip(hillinessIndex_mPerKm / 60, 0, 1)];
score = 100 * mean(components(isfinite(components)), 'omitnan');
end

function score = localRegenOpportunityIndex(downhillShare, steepDownhillShare, elevationLoss_m, routeDistance_km)
descentPerKm = localSafeDivide(elevationLoss_m, routeDistance_km);
components = [ ...
    localClip(downhillShare / 50, 0, 1), ...
    localClip(steepDownhillShare / 20, 0, 1), ...
    localClip(descentPerKm / 40, 0, 1)];
score = 100 * mean(components(isfinite(components)), 'omitnan');
end

function score = localClimateSeverityIndex(coldShare, hotShare, comfortShare, coldExposure_degCh, hotExposure_degCh)
components = [ ...
    localClip(coldShare / 50, 0, 1), ...
    localClip(hotShare / 50, 0, 1), ...
    localClip((100 - comfortShare) / 100, 0, 1), ...
    localClip((coldExposure_degCh + hotExposure_degCh) / 150, 0, 1)];
score = 100 * mean(components(isfinite(components)), 'omitnan');
end

function score = localOverallEnvironmentSeverityIndex(rows)
candidateNames = ["Route Demand Severity Index", "Grade Severity Index", "Climate Severity Index"];
values = NaN(size(candidateNames));
for iName = 1:numel(candidateNames)
    values(iName) = localKPIValue(rows, candidateNames(iName));
end
if all(~isfinite(values))
    score = NaN;
else
    score = mean(values(isfinite(values)), 'omitnan');
end
end

function exposure_degCh = localExposureIntegral(temperature_C, t, reference_C, direction)
cumulative = localCumulativeExposure(temperature_C, t, reference_C, direction);
if isempty(cumulative)
    exposure_degCh = NaN;
else
    exposure_degCh = cumulative(end);
end
end

function cumulative_degCh = localCumulativeExposure(temperature_C, t, reference_C, direction)
temperature_C = double(temperature_C(:));
t = double(t(:));
commonLength = min(numel(temperature_C), numel(t));
temperature_C = temperature_C(1:commonLength);
t = t(1:commonLength);
dt_h = localSampleDurations(t) / 3600;
if direction == "below"
    severity = max(reference_C - temperature_C, 0);
else
    severity = max(temperature_C - reference_C, 0);
end
severity(~isfinite(severity)) = 0;
dt_h(~isfinite(dt_h)) = 0;
cumulative_degCh = cumsum(severity .* dt_h);
end

function [timelineClass, classLabels] = localEnvironmentTimelineClass(desiredSpeed, desiredAccel, roadSlope, ambientTemp, config)
commonLength = max([numel(desiredSpeed), numel(desiredAccel), numel(roadSlope), numel(ambientTemp)]);
timelineClass = NaN(commonLength, 1);
classLabels = ["Stop/low demand", "Cruise or mild route", "Transient demand", "Uphill load", "Downhill regen opportunity", "Thermal stress"];

desiredSpeed = localPadVector(desiredSpeed, commonLength);
desiredAccel = localPadVector(desiredAccel, commonLength);
roadSlope = localPadVector(roadSlope, commonLength);
ambientTemp = localPadVector(ambientTemp, commonLength);

timelineClass(:) = 2;
timelineClass(isfinite(desiredSpeed) & desiredSpeed <= config.Thresholds.CreepSpeed_kmh) = 1;
timelineClass(isfinite(desiredAccel) & abs(desiredAccel) >= config.Thresholds.SignificantAccel_mps2) = 3;
timelineClass(isfinite(roadSlope) & roadSlope > config.Thresholds.SteepSlope_pct) = 4;
timelineClass(isfinite(roadSlope) & roadSlope < -config.Thresholds.SteepSlope_pct) = 5;
timelineClass(isfinite(ambientTemp) & (ambientTemp <= config.Thresholds.ColdAmbient_C | ambientTemp >= config.Thresholds.HotAmbient_C)) = 6;
end

function values = localPadVector(values, targetLength)
values = double(values(:));
if numel(values) < targetLength
    values(end + 1:targetLength, 1) = NaN;
else
    values = values(1:targetLength);
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

function line = localEnvironmentClassification(~, hasDemand, hasGrade, hasAmbient, rows, ~)
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
