function result = Analyze_Driver(analysisData, outputPaths, config)
% Analyze_Driver  PI and feedforward driver-controller behaviour analysis.

result = localInitResult("DRIVER", {'acc_pdl', 'brk_pdl'}, ...
    {'veh_des_vel', 'veh_vel', 'road_slp', 'gr_num', 'emot1_max_av_trq', 'emot2_max_av_trq'});

d = analysisData.Derived;
t = d.time_s(:);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Driver analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.SummaryText = summary;
    result.Suggestions = RCA_MakeSuggestionTable("Driver", recs, evidence);
    return;
end

dt = d.dt_s(:);
vehSpeed = d.vehVel_kmh(:);
desiredSpeed = d.speedDemand_kmh(:);
speedError = d.speedError_kmh(:);
roadSlope = d.roadSlope_pct(:);
accPedal = d.accPedal_pct(:);
brkPedal = d.brkPedal_pct(:);
gear = d.gearNumber(:);
torqueDemand = d.torqueDemandTotal_Nm(:);
torqueActual = d.torqueActualTotal_Nm(:);
torqueLimit = d.torquePositiveLimit_Nm(:);

vehicleMass = d.vehicleMass_kg;
movingMask = localMovingMask(vehSpeed, desiredSpeed, config);
trackingMask = isfinite(speedError) & isfinite(vehSpeed) & isfinite(desiredSpeed) & movingMask;
withinBandMask = abs(speedError) <= config.Thresholds.DriverTrackingBand_kmh;
underspeedMask = speedError > config.Thresholds.DriverTrackingBand_kmh;
overspeedMask = speedError < -config.Thresholds.DriverTrackingBand_kmh;

accActiveMask = isfinite(accPedal) & accPedal >= config.Thresholds.DriverPedalActive_pct;
brkActiveMask = isfinite(brkPedal) & brkPedal >= config.Thresholds.DriverPedalActive_pct;
pedalOverlapMask = isfinite(accPedal) & isfinite(brkPedal) & ...
    accPedal >= config.Thresholds.DriverPedalOverlap_pct & ...
    brkPedal >= config.Thresholds.DriverPedalOverlap_pct;
coastMask = isfinite(accPedal) & isfinite(brkPedal) & ...
    accPedal <= config.Thresholds.DriverCoastPedal_pct & ...
    brkPedal <= config.Thresholds.DriverCoastPedal_pct & movingMask;

desiredAccel = localFiniteDiff(desiredSpeed / 3.6, dt);

nearTorqueLimitMask = false(size(t));
limitApplicableMask = false(size(t));
positiveDemand = max(torqueDemand, 0);
positiveActual = max(torqueActual, 0);
positiveShortfall = max(positiveDemand - positiveActual, 0);
validLimit = isfinite(positiveDemand) & isfinite(torqueLimit) & torqueLimit > 0 & positiveDemand > 0;
limitApplicableMask(validLimit) = true;
nearTorqueLimitMask(validLimit) = positiveDemand(validLimit) >= config.Thresholds.LimitUsageFraction .* torqueLimit(validLimit);

shiftWindowMask = false(size(t));
shiftIndex = find(abs(diff(gear)) > 0 & ~isnan(diff(gear))) + 1;
for iShift = 1:numel(shiftIndex)
    shiftTime = t(shiftIndex(iShift));
    shiftWindowMask = shiftWindowMask | (t >= shiftTime & t <= shiftTime + config.Thresholds.DriverShiftInfluenceWindow_s);
end

if isfinite(vehicleMass)
    rows = RCA_AddKPI(rows, 'Vehicle Mass Context', vehicleMass, 'kg', 'Context', 'Driver', ...
        'VDy_mec_massVehicle_kg', ...
        'Feedforward interpretation uses the logged/spec vehicle mass context when available.');
end

if any(trackingMask)
    trackMae = mean(abs(speedError(trackingMask)), 'omitnan');
    trackRmse = sqrt(mean(speedError(trackingMask).^2, 'omitnan'));
    trackP95 = RCA_Percentile(abs(speedError(trackingMask)), 95);
    trackWithinBand = 100 * RCA_FractionTrue(withinBandMask, trackingMask);
    underspeedShare = 100 * RCA_FractionTrue(underspeedMask, trackingMask);
    overspeedShare = 100 * RCA_FractionTrue(overspeedMask, trackingMask);

    rows = RCA_AddKPI(rows, 'Driver Tracking MAE', trackMae, 'km/h', 'Tracking', 'Driver', ...
        'veh_des_vel + veh_vel', 'Mean absolute speed error over active driving samples.');
    rows = RCA_AddKPI(rows, 'Driver Tracking RMSE', trackRmse, 'km/h', 'Tracking', 'Driver', ...
        'veh_des_vel + veh_vel', 'Root-mean-square speed error highlights occasional large misses.');
    rows = RCA_AddKPI(rows, 'Driver Tracking 95th Percentile', trackP95, 'km/h', 'Tracking', 'Driver', ...
        'veh_des_vel + veh_vel', 'Tail tracking error is useful for stakeholder worst-case review.');
    rows = RCA_AddKPI(rows, 'Tracking Within Band Share', trackWithinBand, '%', 'Tracking', 'Driver', ...
        'veh_des_vel + veh_vel', sprintf('Band is +/- %.1f km/h and is editable in RCA_Config.', config.Thresholds.DriverTrackingBand_kmh));
    rows = RCA_AddKPI(rows, 'Underspeed Share', underspeedShare, '%', 'Performance', 'Driver', ...
        'veh_des_vel + veh_vel', 'Desired speed exceeds actual speed by more than the configured tracking band.');
    rows = RCA_AddKPI(rows, 'Overspeed Share', overspeedShare, '%', 'Performance', 'Driver', ...
        'veh_des_vel + veh_vel', 'Actual speed exceeds desired speed by more than the configured tracking band.');

    summary(end + 1) = sprintf(['Driver tracking stays within +/- %.1f km/h for %.1f%% of active samples. ', ...
        'Underspeed share is %.1f%% and overspeed share is %.1f%%.'], ...
        config.Thresholds.DriverTrackingBand_kmh, trackWithinBand, underspeedShare, overspeedShare);
end

if any(isfinite(accPedal))
    accActiveShare = 100 * RCA_FractionTrue(accActiveMask, isfinite(accPedal) & movingMask);
    highAccelShare = 100 * RCA_FractionTrue(accPedal >= config.Thresholds.DriverHighAccelPedal_pct, isfinite(accPedal) & movingMask);
    rows = RCA_AddKPI(rows, 'Mean Accelerator Pedal', mean(accPedal, 'omitnan'), '%', 'Demand', 'Driver', ...
        'acc_pdl', 'Average accelerator command from the driver controller.');
    rows = RCA_AddKPI(rows, 'Accelerator Pedal 95th Percentile', RCA_Percentile(accPedal, 95), '%', 'Demand', 'Driver', ...
        'acc_pdl', 'High values indicate aggressive or saturation-prone propulsion requests.');
    rows = RCA_AddKPI(rows, 'Accelerator Active Share', accActiveShare, '%', 'Demand', 'Driver', ...
        'acc_pdl', sprintf('Pedal >= %.1f%% is treated as active demand.', config.Thresholds.DriverPedalActive_pct));
    rows = RCA_AddKPI(rows, 'High Accelerator Demand Share', highAccelShare, '%', 'Demand', 'Driver', ...
        'acc_pdl', sprintf('Pedal >= %.1f%% is treated as a strong propulsion request.', config.Thresholds.DriverHighAccelPedal_pct));
end

if any(isfinite(brkPedal))
    brkActiveShare = 100 * RCA_FractionTrue(brkActiveMask, isfinite(brkPedal) & movingMask);
    hardBrakeShare = 100 * RCA_FractionTrue(brkPedal >= config.Thresholds.DriverHighBrakePedal_pct, isfinite(brkPedal) & movingMask);
    rows = RCA_AddKPI(rows, 'Mean Brake Pedal', mean(brkPedal, 'omitnan'), '%', 'Demand', 'Driver', ...
        'brk_pdl', 'Average brake request from the driver controller.');
    rows = RCA_AddKPI(rows, 'Brake Pedal 95th Percentile', RCA_Percentile(brkPedal, 95), '%', 'Demand', 'Driver', ...
        'brk_pdl', 'High values indicate strong or frequent brake intervention.');
    rows = RCA_AddKPI(rows, 'Brake Active Share', brkActiveShare, '%', 'Demand', 'Driver', ...
        'brk_pdl', sprintf('Pedal >= %.1f%% is treated as active braking.', config.Thresholds.DriverPedalActive_pct));
    rows = RCA_AddKPI(rows, 'High Brake Demand Share', hardBrakeShare, '%', 'Demand', 'Driver', ...
        'brk_pdl', sprintf('Pedal >= %.1f%% is treated as strong braking demand.', config.Thresholds.DriverHighBrakePedal_pct));
end

if any(isfinite(accPedal)) && any(isfinite(brkPedal))
    validPedal = movingMask & isfinite(accPedal) & isfinite(brkPedal);
    overlapShare = 100 * RCA_FractionTrue(pedalOverlapMask, validPedal);
    coastShare = 100 * RCA_FractionTrue(coastMask, validPedal);
    rows = RCA_AddKPI(rows, 'Pedal Overlap Share', overlapShare, '%', 'Operation', 'Driver', ...
        'acc_pdl + brk_pdl', ...
        sprintf('Simultaneous pedal demand above %.1f%% is a heuristic arbitration-quality indicator.', config.Thresholds.DriverPedalOverlap_pct));
    rows = RCA_AddKPI(rows, 'Coasting Share', coastShare, '%', 'Efficiency', 'Driver', ...
        'acc_pdl + brk_pdl + veh_vel', ...
        sprintf('Both pedals below %.1f%% are treated as coasting demand.', config.Thresholds.DriverCoastPedal_pct));
end

if any(trackingMask & isfinite(roadSlope))
    uphillMask = trackingMask & roadSlope > config.Thresholds.UphillSlope_pct;
    downhillMask = trackingMask & roadSlope < config.Thresholds.DownhillSlope_pct;
    flatMask = trackingMask & abs(roadSlope) <= config.Thresholds.UphillSlope_pct;
    if any(uphillMask)
        rows = RCA_AddKPI(rows, 'Uphill Tracking MAE', mean(abs(speedError(uphillMask)), 'omitnan'), 'km/h', 'Feedforward', 'Driver', ...
            'veh_des_vel + veh_vel + road_slp', 'Highlights whether uphill compensation is adequate.');
    end
    if any(downhillMask)
        rows = RCA_AddKPI(rows, 'Downhill Overspeed Share', 100 * RCA_FractionTrue(overspeedMask, downhillMask), '%', 'Feedforward', 'Driver', ...
            'veh_des_vel + veh_vel + road_slp', 'Shows whether downhill compensation and braking request are sufficient.');
    end
    if any(flatMask)
        rows = RCA_AddKPI(rows, 'Flat-Road Mean Speed Bias', mean(speedError(flatMask), 'omitnan'), 'km/h', 'Feedforward', 'Driver', ...
            'veh_des_vel + veh_vel + road_slp', 'Flat-road bias helps separate PI tuning issues from route-load effects.');
    end
end

poorTrackingMask = trackingMask & abs(speedError) > config.Thresholds.DriverTrackingBand_kmh;
if any(limitApplicableMask & underspeedMask)
    limitDrivenUnderspeed = 100 * RCA_FractionTrue(nearTorqueLimitMask, trackingMask & underspeedMask & limitApplicableMask);
    rows = RCA_AddKPI(rows, 'Near-Limit Underspeed Share', limitDrivenUnderspeed, '%', 'RootCause', 'Driver', ...
        'veh_des_vel + veh_vel + max available torque', ...
        'Quantifies how much underspeed coincides with near-limit drive-torque capability.');
end

if any(poorTrackingMask)
    shiftPoorShare = 100 * RCA_FractionTrue(shiftWindowMask, poorTrackingMask);
    rows = RCA_AddKPI(rows, 'Shift-Influenced Poor Tracking Share', shiftPoorShare, '%', 'RootCause', 'Driver', ...
        'veh_des_vel + veh_vel + gr_num', ...
        sprintf('Poor tracking inside %.1f s after a gear shift is treated as shift-influenced.', config.Thresholds.DriverShiftInfluenceWindow_s));
end

[eventTable, sampleEventType] = localBuildDriverEventTable(d, desiredAccel, nearTorqueLimitMask, limitApplicableMask, positiveShortfall, config);
dynamicEventMask = height(eventTable) > 0 & ismember(eventTable.EventType, ["Acceleration", "Braking", "Cruise"]) & ...
    eventTable.Duration_s >= config.Thresholds.MinEventDuration_s;
if any(dynamicEventMask)
    dynamicEvents = eventTable(dynamicEventMask, :);
    totalDynamicTime = sum(dynamicEvents.Duration_s, 'omitnan');
    eventTypes = ["Acceleration", "Braking", "Cruise"];
    for iType = 1:numel(eventTypes)
        typeMask = dynamicEvents.EventType == eventTypes(iType);
        eventCount = sum(typeMask);
        timeShare = 100 * sum(dynamicEvents.Duration_s(typeMask), 'omitnan') / max(totalDynamicTime, eps);
        meanDuration = mean(dynamicEvents.Duration_s(typeMask), 'omitnan');
        meanAbsError = mean(dynamicEvents.MeanAbsSpeedError_kmh(typeMask), 'omitnan');
        poorShare = 100 * RCA_FractionTrue(dynamicEvents.Severity >= config.Thresholds.DriverBadEventSeverity, typeMask);

        rows = RCA_AddKPI(rows, string(eventTypes(iType)) + " Event Count", eventCount, 'count', 'Events', 'Driver', ...
            'event segmentation from speed, accel, and pedal behaviour', 'Count of contiguous dynamic events after event filtering.');
        rows = RCA_AddKPI(rows, string(eventTypes(iType)) + " Event Time Share", timeShare, '%', 'Events', 'Driver', ...
            'event segmentation from speed, accel, and pedal behaviour', 'Share of dynamic driving time spent in this event type.');
        rows = RCA_AddKPI(rows, string(eventTypes(iType)) + " Mean Duration", meanDuration, 's', 'Events', 'Driver', ...
            'event segmentation from time base', 'Average duration of this event type.');
        rows = RCA_AddKPI(rows, string(eventTypes(iType)) + " Mean Abs Tracking Error", meanAbsError, 'km/h', 'Events', 'Driver', ...
            'veh_des_vel + veh_vel', 'Average event-level absolute tracking error.');
        rows = RCA_AddKPI(rows, string(eventTypes(iType)) + " Poor Event Share", poorShare, '%', 'Events', 'Driver', ...
            'event severity ranking', sprintf('Event severity >= %.2f is treated as a bad driver event.', config.Thresholds.DriverBadEventSeverity));

        if eventTypes(iType) == "Cruise"
            oscillatoryShare = 100 * RCA_FractionTrue(dynamicEvents.ErrorSignChangeRate_per_s >= config.Thresholds.DriverOscillationRate_per_s, typeMask);
            rows = RCA_AddKPI(rows, 'Cruise Oscillatory Event Share', oscillatoryShare, '%', 'Events', 'Driver', ...
                'veh_des_vel + veh_vel', sprintf('Cruise sign-change rate above %.2f 1/s is treated as oscillatory.', config.Thresholds.DriverOscillationRate_per_s));
        elseif eventTypes(iType) == "Acceleration"
            rows = RCA_AddKPI(rows, 'Acceleration Near-Limit Share', mean(dynamicEvents.NearTorqueLimitShare_pct(typeMask), 'omitnan'), '%', ...
                'Events', 'Driver', 'max available torque + demand torque', ...
                'High values indicate poor acceleration events coincide with drive-torque saturation.');
        elseif eventTypes(iType) == "Braking"
            rows = RCA_AddKPI(rows, 'Braking Overspeed Share', mean(dynamicEvents.OverspeedShare_pct(typeMask), 'omitnan'), '%', ...
                'Events', 'Driver', 'veh_des_vel + veh_vel + brk_pdl', ...
                'Shows how often braking events remain above the desired speed trajectory.');
        end
    end

    badEvents = dynamicEvents(dynamicEvents.Severity >= config.Thresholds.DriverBadEventSeverity, :);
    if ~isempty(badEvents)
        badEvents = sortrows(badEvents, {'Severity', 'MeanAbsSpeedError_kmh'}, {'descend', 'descend'});
        for iBad = 1:min(3, height(badEvents))
            note = sprintf(['Event %d (%s, %.1f to %.1f s). Cause: %s. ', ...
                'Tuning hint: %s'], ...
                badEvents.EventID(iBad), badEvents.EventType(iBad), badEvents.StartTime_s(iBad), badEvents.EndTime_s(iBad), ...
                badEvents.LikelyCause(iBad), badEvents.TuningHint(iBad));
            rows = RCA_AddKPI(rows, sprintf('Bad Driver Event %d Severity', iBad), badEvents.Severity(iBad), '-', ...
                'RootCause', 'Driver', sprintf('Driver event %d', badEvents.EventID(iBad)), note);
            summary(end + 1) = sprintf(['Bad driver event %d is a %s event from %.1f s to %.1f s. ', ...
                'Mean abs speed error is %.2f km/h and the likely cause is %s.'], ...
                badEvents.EventID(iBad), lower(char(badEvents.EventType(iBad))), ...
                badEvents.StartTime_s(iBad), badEvents.EndTime_s(iBad), ...
                badEvents.MeanAbsSpeedError_kmh(iBad), badEvents.LikelyCause(iBad));
        end
    end
else
    badEvents = localEmptyDriverEventTable();
end

if ~any(isfinite(roadSlope))
    result.Warnings(end + 1) = "Road slope is unavailable, so slope feedforward interpretation is limited.";
end
if ~any(isfinite(gear))
    result.Warnings(end + 1) = "Current gear is unavailable, so shift interaction attribution is limited.";
end
if ~any(isfinite(torqueLimit))
    result.Warnings(end + 1) = "Max available drive torque is unavailable, so torque-limit attribution is limited.";
end
if ~isfinite(vehicleMass)
    result.Warnings(end + 1) = "Vehicle mass context is unavailable, so load-feedforward interpretation is limited.";
end

flatCruiseMask = trackingMask & abs(roadSlope) <= config.Thresholds.UphillSlope_pct & sampleEventType == "Cruise" & ~nearTorqueLimitMask;
if any(flatCruiseMask)
    flatCruiseBias = mean(speedError(flatCruiseMask), 'omitnan');
else
    flatCruiseBias = NaN;
end

uphillDemandMask = trackingMask & roadSlope > config.Thresholds.UphillSlope_pct & underspeedMask & accActiveMask;
downhillBrakeMask = trackingMask & roadSlope < config.Thresholds.DownhillSlope_pct & overspeedMask;

if any(limitApplicableMask & trackingMask & underspeedMask)
    limitDrivenUnderspeed = 100 * RCA_FractionTrue(nearTorqueLimitMask, limitApplicableMask & trackingMask & underspeedMask);
else
    limitDrivenUnderspeed = NaN;
end

if isfinite(limitDrivenUnderspeed) && limitDrivenUnderspeed > 40
    recs(end + 1) = "Do not retune PI gains first; a large share of underspeed occurs when requested drive torque is already near the available limit.";
    evidence(end + 1) = sprintf('%.1f%% of underspeed samples coincide with near-limit drive torque.', limitDrivenUnderspeed);
end

if any(uphillDemandMask) && mean(abs(speedError(uphillDemandMask)), 'omitnan') > config.Thresholds.PoorTrackingError_kmh && ...
        (~isfinite(limitDrivenUnderspeed) || limitDrivenUnderspeed < 40)
    recs(end + 1) = "Increase slope/load feedforward before making large PI changes; uphill tracking shortfall appears stronger than flat-road tracking shortfall.";
    evidence(end + 1) = sprintf('Uphill mean absolute speed error is %.2f km/h.', mean(abs(speedError(uphillDemandMask)), 'omitnan'));
end

if isfinite(flatCruiseBias) && abs(flatCruiseBias) > config.Thresholds.DriverTrackingBand_kmh
    if flatCruiseBias > 0
        recs(end + 1) = "Increase integral action modestly or reduce speed-error deadband; the controller shows a persistent flat-road underspeed bias.";
        evidence(end + 1) = sprintf('Flat-road cruise mean speed error is +%.2f km/h.', flatCruiseBias);
    else
        recs(end + 1) = "Reduce steady-state bias on flat-road cruise by trimming feedforward or integral action; the controller tends to run above target speed.";
        evidence(end + 1) = sprintf('Flat-road cruise mean speed error is %.2f km/h.', flatCruiseBias);
    end
end

if any(dynamicEventMask)
    cruiseEvents = eventTable(dynamicEventMask & eventTable.EventType == "Cruise", :);
    cruiseOscillatoryShare = 100 * RCA_FractionTrue(cruiseEvents.ErrorSignChangeRate_per_s >= config.Thresholds.DriverOscillationRate_per_s, true(height(cruiseEvents), 1));
else
    cruiseOscillatoryShare = NaN;
end

if isfinite(cruiseOscillatoryShare) && cruiseOscillatoryShare > 20
    recs(end + 1) = "Reduce proportional aggressiveness or add hysteresis/filtering in cruise; oscillatory sign changes suggest an over-reactive PI loop.";
    evidence(end + 1) = sprintf('%.1f%% of cruise events are oscillatory by the configured sign-change criterion.', cruiseOscillatoryShare);
end

if any(downhillBrakeMask) && mean(brkPedal(downhillBrakeMask), 'omitnan') >= config.Thresholds.DriverHighBrakePedal_pct
    recs(end + 1) = "Strengthen downhill braking feedforward and review regen/friction brake blending; overspeed persists even with strong brake demand.";
    evidence(end + 1) = sprintf('Downhill overspeed coincides with mean brake demand of %.1f%%.', mean(brkPedal(downhillBrakeMask), 'omitnan'));
elseif any(downhillBrakeMask)
    recs(end + 1) = "Increase braking response to downhill overspeed or reduce brake deadband; braking demand appears soft during negative grade events.";
    evidence(end + 1) = sprintf('Downhill overspeed share is %.1f%% of downhill tracking samples.', 100 * RCA_FractionTrue(overspeedMask, trackingMask & roadSlope < config.Thresholds.DownhillSlope_pct));
end

if any(pedalOverlapMask & movingMask) && 100 * RCA_FractionTrue(pedalOverlapMask, movingMask & isfinite(accPedal) & isfinite(brkPedal)) > 1
    recs(end + 1) = "Clean up drive-to-brake arbitration; simultaneous accelerator and brake demand suggests controller handover is not well separated.";
    evidence(end + 1) = sprintf('Pedal overlap occurs during %.2f%% of moving samples.', 100 * RCA_FractionTrue(pedalOverlapMask, movingMask & isfinite(accPedal) & isfinite(brkPedal)));
end

if any(poorTrackingMask) && 100 * RCA_FractionTrue(shiftWindowMask, poorTrackingMask) > 20
    recs(end + 1) = "Add gear-aware feedforward or smoother torque handover through shifts; poor tracking clusters near gear changes.";
    evidence(end + 1) = sprintf('%.1f%% of poor-tracking samples fall inside the configured post-shift window.', 100 * RCA_FractionTrue(shiftWindowMask, poorTrackingMask));
end

driverTableFolder = outputPaths.Tables;
localSafeWriteTable(eventTable, fullfile(driverTableFolder, 'Driver_EventSummary.csv'));
localSafeWriteTable(badEvents, fullfile(driverTableFolder, 'Driver_BadEvents.csv'));

if any(isfinite(desiredSpeed)) || any(isfinite(vehSpeed)) || any(isfinite(accPedal)) || any(isfinite(brkPedal))
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(2, 2, 1);
    if any(isfinite(vehSpeed))
        plot(t, vehSpeed, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    end
    if any(isfinite(desiredSpeed))
        plot(t, desiredSpeed, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    end
    title('Driver Demand Versus Vehicle Speed');
    xlabel('Time (s)');
    ylabel('Speed (km/h)');
    speedLegend = localLegendEntries(any(isfinite(vehSpeed)), any(isfinite(desiredSpeed)));
    if ~isempty(speedLegend)
        legend(speedLegend, 'Location', 'best');
    end
    grid on;

    subplot(2, 2, 2);
    yyaxis left;
    plot(t, speedError, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
    yline(config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', config.Plot.Colors.Neutral);
    yline(-config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', config.Plot.Colors.Neutral);
    ylabel('Speed error (km/h)');
    yyaxis right;
    if any(isfinite(roadSlope))
        plot(t, roadSlope, 'Color', config.Plot.Colors.Slope, 'LineWidth', max(config.Plot.LineWidth - 0.2, 1.0));
        ylabel('Road slope (%)');
    else
        ylabel('Road slope (%)');
    end
    xlabel('Time (s)');
    title('Tracking Error and Route Grade Context');
    grid on;

    subplot(2, 2, 3);
    legendEntries = {};
    if any(isfinite(accPedal))
        plot(t, accPedal, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
        legendEntries{end + 1} = 'Accelerator';
    end
    if any(isfinite(brkPedal))
        plot(t, brkPedal, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
        legendEntries{end + 1} = 'Brake';
    end
    title('Driver Pedal Commands');
    xlabel('Time (s)');
    ylabel('Pedal (%)');
    if ~isempty(legendEntries)
        legend(legendEntries, 'Location', 'best');
    end
    grid on;

    subplot(2, 2, 4);
    if any(isfinite(torqueDemand)) || any(isfinite(torqueLimit))
        if any(isfinite(torqueDemand))
            plot(t, positiveDemand, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
        end
        if any(isfinite(torqueActual))
            plot(t, positiveActual, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
        end
        if any(isfinite(torqueLimit))
            plot(t, torqueLimit, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
        end
        xlabel('Time (s)');
        ylabel('Drive torque (Nm)');
        title('Drive Torque Request, Delivery, and Limit');
        torqueLegend = localTorqueLegend(any(isfinite(torqueDemand)), any(isfinite(torqueActual)), any(isfinite(torqueLimit)));
        if ~isempty(torqueLegend)
            legend(torqueLegend, 'Location', 'best');
        end
    elseif any(isfinite(gear))
        stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
        xlabel('Time (s)');
        ylabel('Gear');
        title('Current Operating Gear');
    else
        text(0.1, 0.5, 'Torque limit and gear context unavailable.', 'Units', 'normalized');
    end
    grid on;
    plotFiles(end + 1) = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Driver'), 'Driver_Controller_Overview', config));
    close(fig);
end

if any(dynamicEventMask)
    dynamicEvents = eventTable(dynamicEventMask, :);
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

    subplot(2, 2, 1);
    stairs(t, localEventCode(sampleEventType), 'Color', config.Plot.Colors.Neutral, 'LineWidth', config.Plot.LineWidth);
    yticks([0 1 2 3 4]);
    yticklabels({'Stop', 'Cruise', 'Acceleration', 'Braking', 'Transition'});
    xlabel('Time (s)');
    ylabel('Event type');
    title('Driver Dynamic Event Segmentation');
    localAddCriteriaBox(gca, sprintf(['Acceleration: accel > %.2f m/s^2 or positive error with accelerator\n', ...
        'Braking: accel < %.2f m/s^2 or brake demand active\n', ...
        'Cruise: low accel and low error transients'], ...
        0.5 * config.Thresholds.SignificantAccel_mps2, 0.5 * config.Thresholds.SignificantDecel_mps2), config);
    grid on;

    subplot(2, 2, 2);
    eventTypes = ["Acceleration", "Braking", "Cruise"];
    counts = zeros(size(eventTypes));
    shares = zeros(size(eventTypes));
    totalDynamicTime = max(sum(dynamicEvents.Duration_s, 'omitnan'), eps);
    for iType = 1:numel(eventTypes)
        typeMask = dynamicEvents.EventType == eventTypes(iType);
        counts(iType) = sum(typeMask);
        shares(iType) = 100 * sum(dynamicEvents.Duration_s(typeMask), 'omitnan') / totalDynamicTime;
    end
    xEvent = 1:numel(eventTypes);
    yyaxis left;
    bar(xEvent, counts, 'FaceColor', config.Plot.Colors.Vehicle);
    ylabel('Count');
    yyaxis right;
    plot(xEvent, shares, '-o', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    ylabel('Time share (%)');
    xticks(xEvent);
    xticklabels(cellstr(eventTypes));
    title('Driver Event Counts and Time Share');
    grid on;

    subplot(2, 2, 3);
    hold on;
    localScatterEvents(dynamicEvents(dynamicEvents.EventType == "Acceleration", :), config.Plot.Colors.Demand);
    localScatterEvents(dynamicEvents(dynamicEvents.EventType == "Braking", :), config.Plot.Colors.Warning);
    localScatterEvents(dynamicEvents(dynamicEvents.EventType == "Cruise", :), config.Plot.Colors.Vehicle);
    xlabel('Mean road slope (%)');
    ylabel('Mean abs speed error (km/h)');
    title('Event Error Versus Route Severity');
    legend({'Acceleration', 'Braking', 'Cruise'}, 'Location', 'best');
    grid on;

    subplot(2, 2, 4);
    rankingTable = dynamicEvents(dynamicEvents.Severity >= config.Thresholds.DriverBadEventSeverity, :);
    if ~isempty(rankingTable)
        rankingTable = sortrows(rankingTable, {'Severity', 'MeanAbsSpeedError_kmh'}, {'descend', 'descend'});
        rankingTable = rankingTable(1:min(5, height(rankingTable)), :);
        barh(rankingTable.Severity, 'FaceColor', config.Plot.Colors.Warning);
        set(gca, 'YDir', 'reverse');
        rankingLabels = strings(height(rankingTable), 1);
        for iRow = 1:height(rankingTable)
            rankingLabels(iRow) = sprintf('E%d %s', rankingTable.EventID(iRow), rankingTable.EventType(iRow));
        end
        yticks(1:height(rankingTable));
        set(gca, 'YTickLabel', cellstr(rankingLabels));
        xlabel('Severity');
        title('Worst Driver Events');
        localAddCriteriaBox(gca, sprintf('Bad-event threshold: severity >= %.2f', config.Thresholds.DriverBadEventSeverity), config);
    else
        text(0.1, 0.5, 'No bad driver events exceeded the configured severity threshold.', 'Units', 'normalized');
    end
    grid on;

    plotFiles(end + 1) = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Driver'), 'Driver_Event_Analysis', config));
    close(fig);
end

if exist('badEvents', 'var') && ~isempty(badEvents)
    dashboardEvents = badEvents(1:min(3, height(badEvents)), :);
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    for iDash = 1:height(dashboardEvents)
        idx = dashboardEvents.StartIndex(iDash):dashboardEvents.EndIndex(iDash);
        subplot(height(dashboardEvents), 1, iDash);
        plot(t(idx), vehSpeed(idx), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
        plot(t(idx), desiredSpeed(idx), '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
        ylabel('Speed (km/h)');
        title(sprintf('Event %d: %s | Cause: %s', dashboardEvents.EventID(iDash), dashboardEvents.EventType(iDash), dashboardEvents.LikelyCause(iDash)));
        if iDash == height(dashboardEvents)
            xlabel('Time (s)');
        end
        grid on;
    end
    plotFiles(end + 1) = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Driver'), 'Driver_Bad_Event_Dashboard', config));
    close(fig);
end

result.Available = ~isempty(rows);
result.KPITable = RCA_FinalizeKPITable(rows);
result.FigureFiles = plotFiles;
result.SummaryText = unique(summary(summary ~= ""));
result.Suggestions = RCA_MakeSuggestionTable("Driver", recs, evidence);
end

function [eventTable, sampleEventType] = localBuildDriverEventTable(derived, desiredAccel, nearTorqueLimitMask, limitApplicableMask, positiveShortfall, config)
t = derived.time_s(:);
dt = derived.dt_s(:);
vehSpeed = derived.vehVel_kmh(:);
desiredSpeed = derived.speedDemand_kmh(:);
speedError = derived.speedError_kmh(:);
vehAcc = derived.vehicleAcceleration_mps2(:);
roadSlope = derived.roadSlope_pct(:);
accPedal = derived.accPedal_pct(:);
brkPedal = derived.brkPedal_pct(:);
gear = derived.gearNumber(:);
distanceStep = derived.distanceStep_km(:);

sampleEventType = repmat("Transition", numel(t), 1);
if isempty(t)
    eventTable = localEmptyDriverEventTable();
    return;
end

movingMask = localMovingMask(vehSpeed, desiredSpeed, config);
accActiveMask = isfinite(accPedal) & accPedal >= config.Thresholds.DriverPedalActive_pct;
brkActiveMask = isfinite(brkPedal) & brkPedal >= config.Thresholds.DriverPedalActive_pct;
trackingBand = config.Thresholds.DriverTrackingBand_kmh;

accelMask = movingMask & ~brkActiveMask & ...
    ((isfinite(desiredAccel) & desiredAccel > 0.5 * config.Thresholds.SignificantAccel_mps2) | ...
    (isfinite(vehAcc) & vehAcc > 0.5 * config.Thresholds.SignificantAccel_mps2) | ...
    (isfinite(speedError) & speedError > trackingBand & accActiveMask));

brakeMask = movingMask & ...
    ((isfinite(desiredAccel) & desiredAccel < 0.5 * config.Thresholds.SignificantDecel_mps2) | ...
    (isfinite(vehAcc) & vehAcc < 0.5 * config.Thresholds.SignificantDecel_mps2) | ...
    brkActiveMask | ...
    (isfinite(speedError) & speedError < -trackingBand));

cruiseMask = movingMask & ~accelMask & ~brakeMask & ...
    (abs(desiredAccel) <= config.Thresholds.CruiseAccelAbs_mps2 | ~isfinite(desiredAccel)) & ...
    (abs(vehAcc) <= max(config.Thresholds.CruiseAccelAbs_mps2, 0.5 * abs(config.Thresholds.SignificantAccel_mps2)) | ~isfinite(vehAcc));

sampleEventType(~movingMask) = "Stop";
sampleEventType(cruiseMask) = "Cruise";
sampleEventType(accelMask) = "Acceleration";
sampleEventType(brakeMask) = "Braking";

startIdx = [1; find(sampleEventType(2:end) ~= sampleEventType(1:end - 1)) + 1];
endIdx = [startIdx(2:end) - 1; numel(t)];
rows = cell(0, 32);

for iEvent = 1:numel(startIdx)
    idx = startIdx(iEvent):endIdx(iEvent);
    duration_s = max(sum(max(dt(idx), 0), 'omitnan'), t(endIdx(iEvent)) - t(startIdx(iEvent)));
    validTrack = isfinite(speedError(idx));
    withinBand = abs(speedError(idx)) <= trackingBand;
    underspeed = speedError(idx) > trackingBand;
    overspeed = speedError(idx) < -trackingBand;
    pedalOverlap = isfinite(accPedal(idx)) & isfinite(brkPedal(idx)) & ...
        accPedal(idx) >= config.Thresholds.DriverPedalOverlap_pct & ...
        brkPedal(idx) >= config.Thresholds.DriverPedalOverlap_pct;
    eventCoastMask = isfinite(accPedal(idx)) & isfinite(brkPedal(idx)) & ...
        accPedal(idx) <= config.Thresholds.DriverCoastPedal_pct & ...
        brkPedal(idx) <= config.Thresholds.DriverCoastPedal_pct;
    shiftCount = sum(abs(diff(gear(idx))) > 0 & ~isnan(diff(gear(idx))));

    metrics = struct();
    metrics.EventType = sampleEventType(startIdx(iEvent));
    metrics.MeanSpeedError_kmh = mean(speedError(idx), 'omitnan');
    metrics.MeanAbsSpeedError_kmh = mean(abs(speedError(idx)), 'omitnan');
    metrics.P95AbsSpeedError_kmh = RCA_Percentile(abs(speedError(idx)), 95);
    metrics.WithinBandShare_pct = 100 * RCA_FractionTrue(withinBand, validTrack);
    metrics.UnderspeedShare_pct = 100 * RCA_FractionTrue(underspeed, validTrack);
    metrics.OverspeedShare_pct = 100 * RCA_FractionTrue(overspeed, validTrack);
    metrics.MeanDesiredAccel_mps2 = mean(desiredAccel(idx), 'omitnan');
    metrics.MeanActualAccel_mps2 = mean(vehAcc(idx), 'omitnan');
    metrics.MeanAccelPedal_pct = mean(accPedal(idx), 'omitnan');
    metrics.MeanBrakePedal_pct = mean(brkPedal(idx), 'omitnan');
    metrics.PedalOverlapShare_pct = 100 * RCA_FractionTrue(pedalOverlap, isfinite(accPedal(idx)) & isfinite(brkPedal(idx)));
    metrics.CoastShare_pct = 100 * RCA_FractionTrue(eventCoastMask, isfinite(accPedal(idx)) & isfinite(brkPedal(idx)));
    metrics.MeanSlope_pct = mean(roadSlope(idx), 'omitnan');
    metrics.ShiftCount = shiftCount;
    metrics.NearTorqueLimitShare_pct = 100 * RCA_FractionTrue(nearTorqueLimitMask(idx), limitApplicableMask(idx));
    metrics.TorqueShortfallMAE_Nm = mean(positiveShortfall(idx), 'omitnan');
    metrics.ErrorSignChangeRate_per_s = localErrorSignChangeRate(speedError(idx), trackingBand, duration_s);

    severity = localEventSeverity(metrics, config);
    [likelyCause, confidenceNote, tuningHint] = localInterpretDriverEvent(metrics, config);

    rows(end + 1, :) = {iEvent, startIdx(iEvent), endIdx(iEvent), ...
        t(startIdx(iEvent)), t(endIdx(iEvent)), duration_s, sampleEventType(startIdx(iEvent)), ...
        sum(distanceStep(idx), 'omitnan'), mean(desiredSpeed(idx), 'omitnan'), mean(vehSpeed(idx), 'omitnan'), ...
        metrics.MeanSpeedError_kmh, metrics.MeanAbsSpeedError_kmh, metrics.P95AbsSpeedError_kmh, ...
        metrics.WithinBandShare_pct, metrics.UnderspeedShare_pct, metrics.OverspeedShare_pct, ...
        metrics.MeanDesiredAccel_mps2, metrics.MeanActualAccel_mps2, metrics.MeanAccelPedal_pct, ...
        metrics.MeanBrakePedal_pct, metrics.PedalOverlapShare_pct, metrics.CoastShare_pct, ...
        metrics.MeanSlope_pct, mean(gear(idx), 'omitnan'), metrics.ShiftCount, ...
        metrics.NearTorqueLimitShare_pct, metrics.TorqueShortfallMAE_Nm, metrics.ErrorSignChangeRate_per_s, ...
        severity, string(likelyCause), string(confidenceNote), string(tuningHint)}; %#ok<AGROW>
end

eventTable = cell2table(rows, 'VariableNames', {'EventID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
    'Duration_s', 'EventType', 'Distance_km', 'MeanDesiredSpeed_kmh', 'MeanVehicleSpeed_kmh', ...
    'MeanSpeedError_kmh', 'MeanAbsSpeedError_kmh', 'P95AbsSpeedError_kmh', 'WithinBandShare_pct', ...
    'UnderspeedShare_pct', 'OverspeedShare_pct', 'MeanDesiredAccel_mps2', 'MeanActualAccel_mps2', ...
    'MeanAccelPedal_pct', 'MeanBrakePedal_pct', 'PedalOverlapShare_pct', 'CoastShare_pct', ...
    'MeanSlope_pct', 'MeanGear', 'ShiftCount', 'NearTorqueLimitShare_pct', 'TorqueShortfallMAE_Nm', ...
    'ErrorSignChangeRate_per_s', 'Severity', 'LikelyCause', 'ConfidenceNote', 'TuningHint'});
end

function severity = localEventSeverity(metrics, config)
baseSeverity = metrics.MeanAbsSpeedError_kmh / max(config.Thresholds.PoorTrackingError_kmh, eps);
tailSeverity = metrics.P95AbsSpeedError_kmh / max(config.Thresholds.SevereTrackingError_kmh, eps);

switch metrics.EventType
    case "Acceleration"
        severity = max(baseSeverity, tailSeverity) + ...
            0.35 * metrics.UnderspeedShare_pct / 50 + ...
            0.25 * metrics.NearTorqueLimitShare_pct / 25 + ...
            0.20 * max(metrics.MeanSlope_pct, 0) / max(config.Thresholds.SteepSlope_pct, eps) + ...
            0.15 * min(metrics.ShiftCount, 2);
    case "Braking"
        severity = max(baseSeverity, tailSeverity) + ...
            0.35 * metrics.OverspeedShare_pct / 50 + ...
            0.25 * max(-metrics.MeanSlope_pct, 0) / max(config.Thresholds.SteepSlope_pct, eps) + ...
            0.20 * metrics.PedalOverlapShare_pct / 10 + ...
            0.10 * min(metrics.ShiftCount, 2);
    case "Cruise"
        severity = max([baseSeverity, tailSeverity, abs(metrics.MeanSpeedError_kmh) / max(config.Thresholds.DriverTrackingBand_kmh, eps)]) + ...
            0.35 * metrics.ErrorSignChangeRate_per_s / max(config.Thresholds.DriverOscillationRate_per_s, eps) + ...
            0.15 * metrics.PedalOverlapShare_pct / 10;
    otherwise
        severity = max(baseSeverity, tailSeverity);
end
end

function [likelyCause, confidenceNote, tuningHint] = localInterpretDriverEvent(metrics, config)
likelyCause = "General tracking mismatch";
confidenceNote = "Low confidence because the observed pattern is not strongly diagnostic.";
tuningHint = "Review driver-controller behaviour together with torque availability and route context.";

switch metrics.EventType
    case "Acceleration"
        if metrics.NearTorqueLimitShare_pct >= 40
            likelyCause = "Drive torque availability limit";
            confidenceNote = "Medium-high confidence because acceleration shortfall coincides with near-limit drive torque.";
            tuningHint = "Before increasing PI gains, review torque-limit scheduling, battery power limits, and available-torque assumptions.";
        elseif metrics.ShiftCount > 0 && metrics.MeanAbsSpeedError_kmh > config.Thresholds.PoorTrackingError_kmh
            likelyCause = "Gear-shift interaction";
            confidenceNote = "Medium confidence because poor tracking occurs together with gear transitions.";
            tuningHint = "Add gear-aware feedforward or smoother torque handover through shifts.";
        elseif metrics.MeanSlope_pct > config.Thresholds.UphillSlope_pct && metrics.UnderspeedShare_pct > 40
            likelyCause = "Slope/load feedforward too weak";
            confidenceNote = "Medium confidence because uphill acceleration remains below target without strong saturation evidence.";
            tuningHint = "Increase grade or mass feedforward before making large PI-gain changes.";
        elseif metrics.MeanAccelPedal_pct < config.Thresholds.DriverPedalActive_pct + 5 && metrics.UnderspeedShare_pct > 40
            likelyCause = "PI response too soft or deadband too large";
            confidenceNote = "Medium confidence because positive speed error persists without strong acceleration demand.";
            tuningHint = "Increase proportional response modestly or reduce deadband; if bias remains, increase integral action carefully.";
        else
            likelyCause = "Acceleration transient mismatch";
            confidenceNote = "Medium confidence because the event shows response lag without a single dominant limiter.";
            tuningHint = "Review PI gain balance and feedforward shaping around demand ramps.";
        end

    case "Braking"
        if metrics.MeanSlope_pct < config.Thresholds.DownhillSlope_pct && metrics.OverspeedShare_pct > 40 && ...
                metrics.MeanBrakePedal_pct >= config.Thresholds.DriverHighBrakePedal_pct
            likelyCause = "Downhill braking/feedforward too weak or braking capability limited";
            confidenceNote = "Medium-high confidence because overspeed persists during strong braking demand on negative grade.";
            tuningHint = "Strengthen downhill feedforward and review regen/friction brake blending and brake authority.";
        elseif metrics.MeanSlope_pct < config.Thresholds.DownhillSlope_pct && metrics.OverspeedShare_pct > 40
            likelyCause = "Brake PI response too soft";
            confidenceNote = "Medium confidence because downhill overspeed occurs without strong brake demand.";
            tuningHint = "Increase braking response to negative speed error or reduce brake deadband.";
        elseif metrics.PedalOverlapShare_pct > 2
            likelyCause = "Drive-brake arbitration overlap";
            confidenceNote = "Medium confidence because accelerator and brake commands overlap materially inside the event.";
            tuningHint = "Separate traction release and brake build-up logic more cleanly.";
        else
            likelyCause = "Braking transient mismatch";
            confidenceNote = "Low-medium confidence because braking error exists without a single dominant signature.";
            tuningHint = "Review braking PI calibration together with downhill and regenerative feedforward behaviour.";
        end

    case "Cruise"
        if metrics.ErrorSignChangeRate_per_s >= config.Thresholds.DriverOscillationRate_per_s && ...
                metrics.MeanAbsSpeedError_kmh > config.Thresholds.DriverTrackingBand_kmh
            likelyCause = "Cruise oscillation or over-aggressive PI gains";
            confidenceNote = "Medium-high confidence because the speed error repeatedly changes sign during nominal cruise.";
            tuningHint = "Reduce proportional aggressiveness, add hysteresis/filtering, and verify anti-windup behaviour.";
        elseif abs(metrics.MeanSpeedError_kmh) > config.Thresholds.DriverTrackingBand_kmh && metrics.NearTorqueLimitShare_pct < 20
            likelyCause = "Steady-state PI bias or feedforward bias";
            confidenceNote = "Medium confidence because the event shows sustained offset without strong torque saturation.";
            tuningHint = "Trim feedforward bias and increase integral action cautiously if the same bias persists on flat road.";
        elseif metrics.ShiftCount > 0 && metrics.MeanAbsSpeedError_kmh > config.Thresholds.DriverTrackingBand_kmh
            likelyCause = "Gear-induced cruise disturbance";
            confidenceNote = "Medium confidence because cruise error aligns with gear changes.";
            tuningHint = "Coordinate gear shift and driver-controller torque release/reapply logic.";
        else
            likelyCause = "Cruise tracking acceptable";
            confidenceNote = "High confidence because the event remains largely inside the configured tracking band.";
            tuningHint = "No major controller retuning is indicated by this event.";
        end
end
end

function mask = localMovingMask(vehSpeed, desiredSpeed, config)
mask = (isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh) | ...
    (isfinite(desiredSpeed) & desiredSpeed > config.Thresholds.StopSpeed_kmh);
end

function derivative = localFiniteDiff(signal, dt)
signal = double(signal(:));
dt = double(dt(:));
derivative = NaN(size(signal));
if isempty(signal)
    return;
end
if numel(signal) == 1
    derivative(1) = 0;
    return;
end
derivative(1) = 0;
deltaSignal = diff(signal);
deltaTime = dt(2:end);
deltaTime(~isfinite(deltaTime) | deltaTime <= 0) = NaN;
derivative(2:end) = deltaSignal ./ deltaTime;
end

function rate = localErrorSignChangeRate(speedError, band, duration_s)
signValue = zeros(size(speedError));
signValue(speedError > band) = 1;
signValue(speedError < -band) = -1;
signValue = signValue(signValue ~= 0);
if numel(signValue) <= 1 || ~isfinite(duration_s) || duration_s <= 0
    rate = 0;
    return;
end
rate = sum(signValue(2:end) ~= signValue(1:end - 1)) / duration_s;
end

function code = localEventCode(sampleEventType)
code = zeros(size(sampleEventType));
code(sampleEventType == "Cruise") = 1;
code(sampleEventType == "Acceleration") = 2;
code(sampleEventType == "Braking") = 3;
code(sampleEventType == "Transition") = 4;
end

function entries = localLegendEntries(hasVehicle, hasDemand)
entries = {};
if hasVehicle
    entries{end + 1} = 'Vehicle speed';
end
if hasDemand
    entries{end + 1} = 'Desired speed';
end
end

function entries = localTorqueLegend(hasDemand, hasActual, hasLimit)
entries = {};
if hasDemand
    entries{end + 1} = 'Demand torque';
end
if hasActual
    entries{end + 1} = 'Actual torque';
end
if hasLimit
    entries{end + 1} = 'Positive torque limit';
end
end

function localScatterEvents(eventTable, colorValue)
if isempty(eventTable)
    return;
end
scatter(eventTable.MeanSlope_pct, eventTable.MeanAbsSpeedError_kmh, ...
    max(24, 8 * eventTable.Duration_s), colorValue, 'filled', 'MarkerEdgeColor', [0.2 0.2 0.2]);
end

function localAddCriteriaBox(axisHandle, textValue, config)
text(axisHandle, 0.02, 0.98, textValue, 'Units', 'normalized', 'VerticalAlignment', 'top', ...
    'BackgroundColor', [1 1 1], 'EdgeColor', [0.85 0.85 0.85], 'Margin', 6, ...
    'FontSize', max(config.Plot.FontSize - 1, 8));
end

function tableValue = localEmptyDriverEventTable()
tableValue = cell2table(cell(0, 32), 'VariableNames', {'EventID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
    'Duration_s', 'EventType', 'Distance_km', 'MeanDesiredSpeed_kmh', 'MeanVehicleSpeed_kmh', ...
    'MeanSpeedError_kmh', 'MeanAbsSpeedError_kmh', 'P95AbsSpeedError_kmh', 'WithinBandShare_pct', ...
    'UnderspeedShare_pct', 'OverspeedShare_pct', 'MeanDesiredAccel_mps2', 'MeanActualAccel_mps2', ...
    'MeanAccelPedal_pct', 'MeanBrakePedal_pct', 'PedalOverlapShare_pct', 'CoastShare_pct', ...
    'MeanSlope_pct', 'MeanGear', 'ShiftCount', 'NearTorqueLimitShare_pct', 'TorqueShortfallMAE_Nm', ...
    'ErrorSignChangeRate_per_s', 'Severity', 'LikelyCause', 'ConfidenceNote', 'TuningHint'});
end

function localSafeWriteTable(tableValue, filePath)
try
    writetable(tableValue, filePath);
catch
end
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
