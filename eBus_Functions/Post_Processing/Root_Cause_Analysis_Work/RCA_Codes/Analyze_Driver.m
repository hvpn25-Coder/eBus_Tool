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

summary(end + 1) = "General driver-controller tuning hints are included for acceleration, braking, and cruise phases so calibration discussion is not limited only to the detected bad events.";

recs(end + 1) = "Acceleration phase tuning hint: calibrate positive speed-error response, grade/load feedforward, and torque-request ramp shaping together; avoid increasing PI aggressiveness without checking torque-limit interaction and shift handover quality.";
evidence(end + 1) = "General acceleration guidance based on veh_des_vel, veh_vel, road_slp, gr_num, and max available drive torque.";

recs(end + 1) = "Braking phase tuning hint: tune negative speed-error entry, downhill braking feedforward, and regen/friction-brake coordination together; also review traction-release to brake-build-up arbitration so overspeed is corrected without delay.";
evidence(end + 1) = "General braking guidance based on veh_des_vel, veh_vel, road_slp, and brk_pdl with braking-event RCA context.";

recs(end + 1) = "Cruise phase tuning hint: tune steady-state bias removal with integral action carefully, then use deadband, hysteresis, filtering, and anti-windup to avoid oscillation or hunting around the desired speed.";
evidence(end + 1) = "General cruise guidance based on flat-road cruise bias, speed-error sign-change behaviour, and pedal stability trends.";

driverTableFolder = outputPaths.Tables;
localSafeWriteTable(eventTable, fullfile(driverTableFolder, 'Driver_EventSummary.csv'));
localSafeWriteTable(badEvents, fullfile(driverTableFolder, 'Driver_BadEvents.csv'));
driverBadSegments = localBuildDriverBadSegmentTable(t, desiredSpeed, vehSpeed, speedError, accPedal, brkPedal, gear, roadSlope, config);
localSafeWriteTable(driverBadSegments, fullfile(driverTableFolder, 'Driver_TrackingBadSegments.csv'));
if ~isempty(driverBadSegments)
    badSegmentTimeShare = 100 * sum(driverBadSegments.Duration_s, 'omitnan') / max(t(end) - t(1), eps);
    rows = RCA_AddKPI(rows, 'Driver Bad Segment Count', height(driverBadSegments), 'count', 'RootCause', 'Driver', ...
        'veh_des_vel + veh_vel', ...
        sprintf('|speed error| >= %.1f km/h for at least %.1f s.', config.Thresholds.DriverBadSegmentError_kmh, config.Thresholds.DriverBadSegmentMinDuration_s));
    rows = RCA_AddKPI(rows, 'Driver Bad Segment Time Share', badSegmentTimeShare, '%', 'RootCause', 'Driver', ...
        'veh_des_vel + veh_vel + time', ...
        'Share of trip time occupied by threshold-based driver bad segments.');
    summary(end + 1) = sprintf(['Driver bad-segment detection uses |speed error| >= %.1f km/h for at least %.1f s. ', ...
        '%d bad segments were found, covering %.1f%% of trip time.'], ...
        config.Thresholds.DriverBadSegmentError_kmh, config.Thresholds.DriverBadSegmentMinDuration_s, ...
        height(driverBadSegments), badSegmentTimeShare);
end
driverFigureFolder = fullfile(outputPaths.FiguresSubsystem, 'Driver');
if any(dynamicEventMask)
    dynamicEvents = eventTable(dynamicEventMask, :);
else
    dynamicEvents = localEmptyDriverEventTable();
end

plotFiles = localAppendPlotFile(plotFiles, localPlotDriverOverview(driverFigureFolder, t, desiredSpeed, vehSpeed, speedError, accPedal, brkPedal, gear, roadSlope, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotDriverEventHighlights(driverFigureFolder, t, desiredSpeed, vehSpeed, speedError, dynamicEvents, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotDriverErrorMaps(driverFigureFolder, gear, roadSlope, speedError, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotDriverGearwiseTracking(driverFigureFolder, t, desiredSpeed, vehSpeed, speedError, gear, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotDriverBadSegments(driverFigureFolder, t, desiredSpeed, vehSpeed, speedError, driverBadSegments, config));
plotFiles = [plotFiles; reshape(localPlotDriverWorstSegments(driverFigureFolder, t, desiredSpeed, vehSpeed, speedError, accPedal, brkPedal, gear, roadSlope, driverBadSegments, config), [], 1)]; %#ok<AGROW>
plotFiles = plotFiles(plotFiles ~= "");

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

function tableValue = localEmptyDriverBadSegmentTable()
tableValue = cell2table(cell(0, 13), 'VariableNames', {'DriverSegmentID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
    'Duration_s', 'MaxAbsError_kmh', 'MeanAbsError_kmh', 'MeanDesiredSpeed_kmh', 'MeanVehicleSpeed_kmh', ...
    'DominantGear', 'GradeClass', 'MeanSlope_pct'});
end

function localSafeWriteTable(tableValue, filePath)
try
    writetable(tableValue, filePath);
catch
end
end

function badSegmentTable = localBuildDriverBadSegmentTable(t, desiredSpeed, vehSpeed, speedError, accPedal, brkPedal, gear, roadSlope, config)
badSegmentTable = localEmptyDriverBadSegmentTable();
if isempty(t) || ~any(isfinite(speedError))
    return;
end

errorMask = isfinite(speedError) & abs(speedError) >= config.Thresholds.DriverBadSegmentError_kmh;
startIdx = find(errorMask & ~[false; errorMask(1:end-1)]);
endIdx = find(errorMask & ~[errorMask(2:end); false]);
if isempty(startIdx)
    return;
end

rows = cell(0, 13);
for iSeg = 1:numel(startIdx)
    idx = startIdx(iSeg):endIdx(iSeg);
    duration_s = max(t(endIdx(iSeg)) - t(startIdx(iSeg)), 0);
    if duration_s < config.Thresholds.DriverBadSegmentMinDuration_s
        continue;
    end
    meanSlope = mean(roadSlope(idx), 'omitnan');
    rows(end + 1, :) = { ...
        size(rows, 1) + 1, startIdx(iSeg), endIdx(iSeg), t(startIdx(iSeg)), t(endIdx(iSeg)), duration_s, ...
        max(abs(speedError(idx)), [], 'omitnan'), mean(abs(speedError(idx)), 'omitnan'), ...
        mean(desiredSpeed(idx), 'omitnan'), mean(vehSpeed(idx), 'omitnan'), ...
        localDominantValue(gear(idx)), localGradeClass(meanSlope, config), meanSlope}; %#ok<AGROW>
end

if isempty(rows)
    return;
end

badSegmentTable = cell2table(rows, 'VariableNames', {'DriverSegmentID', 'StartIndex', 'EndIndex', 'StartTime_s', 'EndTime_s', ...
    'Duration_s', 'MaxAbsError_kmh', 'MeanAbsError_kmh', 'MeanDesiredSpeed_kmh', 'MeanVehicleSpeed_kmh', ...
    'DominantGear', 'GradeClass', 'MeanSlope_pct'});
end

function plotFile = localPlotDriverOverview(outputFolder, t, desiredSpeed, vehSpeed, speedError, accPedal, brkPedal, gear, roadSlope, config)
plotFile = "";
if ~(any(isfinite(desiredSpeed)) || any(isfinite(vehSpeed)) || any(isfinite(accPedal)) || any(isfinite(brkPedal)) || any(isfinite(gear)))
    return;
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(4, 1, 1);
hold on;
if any(isfinite(desiredSpeed))
    plot(t, desiredSpeed, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
end
if any(isfinite(vehSpeed))
    plot(t, vehSpeed, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
end
title('Reference vs Actual Speed');
ylabel('Speed [km/h]');
legend(localSpeedLegend(any(isfinite(desiredSpeed)), any(isfinite(vehSpeed))), 'Location', 'best');
grid on;

subplot(4, 1, 2);
plot(t, speedError, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
yline(-config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
ylabel('Error [km/h]');
title('Speed Tracking Error');
grid on;

subplot(4, 1, 3);
hold on;
if any(isfinite(accPedal))
    plot(t, accPedal, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
end
if any(isfinite(brkPedal))
    plot(t, brkPedal, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
end
ylabel('Pedal [%]');
title('Pedal Commands');
legend(localPedalLegend(any(isfinite(accPedal)), any(isfinite(brkPedal))), 'Location', 'best');
grid on;

subplot(4, 1, 4);
yyaxis left;
if any(isfinite(gear))
    stairs(t, gear, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
end
ylabel('Gear [-]');
yyaxis right;
if any(isfinite(roadSlope))
    plot(t, roadSlope, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
end
ylabel('Slope [%]');
xlabel('Time [s]');
title('Gear & Road Slope');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Driver_Tracking_Overview', config));
close(fig);
end

function plotFile = localPlotDriverEventHighlights(outputFolder, t, desiredSpeed, vehSpeed, speedError, eventTable, config)
plotFile = "";
if isempty(eventTable) || height(eventTable) == 0
    return;
end

accelEvents = eventTable(eventTable.EventType == "Acceleration", :);
brakeEvents = eventTable(eventTable.EventType == "Braking", :);
if isempty(accelEvents) && isempty(brakeEvents)
    return;
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 1, 1);
hold on;
refHandle = plot(t, desiredSpeed, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
actHandle = plot(t, vehSpeed, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
accPatch = localShadeIntervals(gca, accelEvents.StartTime_s, accelEvents.EndTime_s, [0.80 0.93 0.80], 0.65);
brkPatch = localShadeIntervals(gca, brakeEvents.StartTime_s, brakeEvents.EndTime_s, [0.82 0.92 0.97], 0.70);
title('Reference vs Actual Speed (Accel/Brake Events Highlighted)');
ylabel('Speed [km/h]');
[eventLegendHandles, eventLegendLabels] = localEventLegend(refHandle, actHandle, accPatch, brkPatch);
legend(eventLegendHandles, eventLegendLabels, 'Location', 'best');
grid on;

subplot(2, 1, 2);
hold on;
plot(t, speedError, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
yline(-config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
localShadeIntervals(gca, accelEvents.StartTime_s, accelEvents.EndTime_s, [0.80 0.93 0.80], 0.65);
localShadeIntervals(gca, brakeEvents.StartTime_s, brakeEvents.EndTime_s, [0.82 0.92 0.97], 0.70);
title('Speed Error (Accel/Brake Events Highlighted)');
ylabel('Error [km/h]');
xlabel('Time [s]');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Driver_Event_Highlights', config));
close(fig);
end

function plotFile = localPlotDriverErrorMaps(outputFolder, gear, roadSlope, speedError, config)
plotFile = "";
valid = isfinite(gear) & isfinite(roadSlope) & isfinite(speedError);
if ~any(valid)
    return;
end

gearValues = unique(gear(valid));
gearValues = gearValues(:)';
slopeEdges = [-Inf, -config.Thresholds.SteepSlope_pct, -config.Thresholds.UphillSlope_pct, ...
    config.Thresholds.UphillSlope_pct, config.Thresholds.SteepSlope_pct, Inf];
slopeLabels = {sprintf('< -%.1f', config.Thresholds.SteepSlope_pct), ...
    sprintf('-%.1f to -%.1f', config.Thresholds.SteepSlope_pct, config.Thresholds.UphillSlope_pct), ...
    sprintf('|slope| <= %.1f', config.Thresholds.UphillSlope_pct), ...
    sprintf('%.1f to %.1f', config.Thresholds.UphillSlope_pct, config.Thresholds.SteepSlope_pct), ...
    sprintf('> %.1f', config.Thresholds.SteepSlope_pct)};

heatmapData = NaN(numel(gearValues), numel(slopeEdges) - 1);
meanErrorPerGear = NaN(numel(gearValues), 1);
for iGear = 1:numel(gearValues)
    gearMask = valid & gear == gearValues(iGear);
    meanErrorPerGear(iGear) = mean(speedError(gearMask), 'omitnan');
    for iBin = 1:(numel(slopeEdges) - 1)
        binMask = gearMask & roadSlope >= slopeEdges(iBin) & roadSlope < slopeEdges(iBin + 1);
        if iBin == numel(slopeEdges) - 1
            binMask = gearMask & roadSlope >= slopeEdges(iBin);
        end
        if any(binMask)
            heatmapData(iGear, iBin) = mean(abs(speedError(binMask)), 'omitnan');
        end
    end
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 2, [1 2]);
imagesc(heatmapData);
set(gca, 'YTick', 1:numel(gearValues), 'YTickLabel', arrayfun(@(x) sprintf('%.0f', x), gearValues, 'UniformOutput', false));
set(gca, 'XTick', 1:numel(slopeLabels), 'XTickLabel', slopeLabels);
xlabel('Slope Range');
ylabel('Gear');
title('Mean Velocity Error Heatmap');
cb = colorbar;
cb.Label.String = 'Mean |speed error| [km/h]';
grid on;

subplot(2, 2, 3);
scatter(roadSlope(valid), speedError(valid), 16, config.Plot.Colors.Vehicle, 'filled');
xlabel('Slope');
ylabel('Velocity Error');
title('Error vs Slope');
grid on;

subplot(2, 2, 4);
bar(gearValues, meanErrorPerGear, 'FaceColor', config.Plot.Colors.Vehicle);
xlabel('Gear');
ylabel('Mean Error');
title('Mean Error per Gear');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Driver_Error_Maps', config));
close(fig);
end

function plotFile = localPlotDriverGearwiseTracking(outputFolder, t, desiredSpeed, vehSpeed, speedError, gear, config)
plotFile = "";
validGear = isfinite(gear);
if ~any(validGear)
    return;
end

gearValues = unique(gear(validGear));
gearValues = gearValues(:)';
fig = figure('Color', 'w', 'Position', [100 100 1300 max(500, 260 * numel(gearValues))]);

for iGear = 1:numel(gearValues)
    mask = gear == gearValues(iGear);
    gearRef = NaN(size(desiredSpeed));
    gearAct = NaN(size(vehSpeed));
    gearErr = NaN(size(speedError));
    gearRef(mask) = desiredSpeed(mask);
    gearAct(mask) = vehSpeed(mask);
    gearErr(mask) = speedError(mask);

    subplot(numel(gearValues), 1, iGear);
    yyaxis left;
    plot(t, gearRef, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    plot(t, gearAct, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    ylabel('Speed [km/h]');
    yyaxis right;
    plot(t, gearErr, 'Color', config.Plot.Colors.Warning, 'LineWidth', max(config.Plot.LineWidth - 0.1, 1.0));
    ylabel('Error [km/h]');
    title(sprintf('Gear %.0f: Speed & Error', gearValues(iGear)));
    if iGear == numel(gearValues)
        xlabel('Time [s]');
    end
    grid on;
end

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Driver_Gearwise_Tracking', config));
close(fig);
end

function plotFile = localPlotDriverBadSegments(outputFolder, t, desiredSpeed, vehSpeed, speedError, badSegmentTable, config)
plotFile = "";
if isempty(badSegmentTable) || height(badSegmentTable) == 0
    return;
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 1, 1);
hold on;
refHandle = plot(t, desiredSpeed, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
actHandle = plot(t, vehSpeed, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
badPatch = localShadeIntervals(gca, badSegmentTable.StartTime_s, badSegmentTable.EndTime_s, [0.97 0.88 0.88], 0.80);
title('Reference vs Actual Speed (Bad Segments Highlighted)');
ylabel('Speed [km/h]');
legend([refHandle, actHandle, badPatch], {'v_{ref}', 'v_{act}', 'Bad segments'}, 'Location', 'best');
grid on;

subplot(2, 1, 2);
hold on;
plot(t, speedError, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
yline(config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
yline(-config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
localShadeIntervals(gca, badSegmentTable.StartTime_s, badSegmentTable.EndTime_s, [0.97 0.88 0.88], 0.80);
title('Speed Error (Bad Segments Highlighted)');
ylabel('Error [km/h]');
xlabel('Time [s]');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'Driver_Bad_Segments', config));
close(fig);
end

function plotFiles = localPlotDriverWorstSegments(outputFolder, t, desiredSpeed, vehSpeed, speedError, accPedal, brkPedal, gear, roadSlope, badSegmentTable, config)
plotFiles = strings(0, 1);
if isempty(badSegmentTable) || height(badSegmentTable) == 0
    return;
end

candidateSegments = sortrows(badSegmentTable, {'MaxAbsError_kmh', 'MeanAbsError_kmh', 'Duration_s'}, {'descend', 'descend', 'descend'});
candidateSegments = candidateSegments(1:min(3, height(candidateSegments)), :);

for iSeg = 1:height(candidateSegments)
    idx = candidateSegments.StartIndex(iSeg):candidateSegments.EndIndex(iSeg);
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

    subplot(4, 1, 1);
    hold on;
    plot(t(idx), desiredSpeed(idx), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
    plot(t(idx), vehSpeed(idx), 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    localShadeIntervals(gca, candidateSegments.StartTime_s(iSeg), candidateSegments.EndTime_s(iSeg), [0.97 0.92 0.92], 0.85);
    ylabel('Speed [km/h]');
    title(sprintf('Segment %d: t=%.1f - %.1f s (dur=%.2f s, max|e|=%.1f km/h)', ...
        candidateSegments.DriverSegmentID(iSeg), candidateSegments.StartTime_s(iSeg), candidateSegments.EndTime_s(iSeg), ...
        candidateSegments.Duration_s(iSeg), candidateSegments.MaxAbsError_kmh(iSeg)));
    legend({'v_{ref}', 'v_{act}'}, 'Location', 'best');
    grid on;

    subplot(4, 1, 2);
    plot(t(idx), speedError(idx), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    yline(config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
    yline(-config.Thresholds.DriverTrackingBand_kmh, '--', 'Color', [0.6 0.6 0.6], 'LineWidth', 1.0);
    ylabel('Error [km/h]');
    title('Speed Error');
    grid on;

    subplot(4, 1, 3);
    hold on;
    plot(t(idx), accPedal(idx), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
    plot(t(idx), brkPedal(idx), 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
    ylabel('Pedal [%]');
    title('Accel / Brake Commands');
    legend({'Accel', 'Brake'}, 'Location', 'best');
    grid on;

    subplot(4, 1, 4);
    yyaxis left;
    stairs(t(idx), gear(idx), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
    ylabel('Gear [-]');
    yyaxis right;
    plot(t(idx), roadSlope(idx), 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
    ylabel('Slope [%]');
    xlabel('Time [s]');
    title(sprintf('Gear & Slope (dominant gear: %.0f, slope: %s)', ...
        candidateSegments.DominantGear(iSeg), candidateSegments.GradeClass(iSeg)));
    grid on;

    plotFiles(end + 1, 1) = string(RCA_SaveFigure(fig, outputFolder, sprintf('Driver_Worst_Segment_%02d', iSeg), config)); %#ok<AGROW>
    close(fig);
end
end

function plotFiles = localAppendPlotFile(plotFiles, plotFile)
if nargin < 2 || strlength(string(plotFile)) == 0
    return;
end
plotFiles(end + 1, 1) = string(plotFile);
end

function patchHandle = localShadeIntervals(axisHandle, startTimes, endTimes, colorValue, alphaValue)
patchHandle = gobjects(1, 1);
startTimes = double(startTimes(:));
endTimes = double(endTimes(:));
if isempty(startTimes) || isempty(endTimes)
    return;
end
valid = isfinite(startTimes) & isfinite(endTimes) & endTimes >= startTimes;
startTimes = startTimes(valid);
endTimes = endTimes(valid);
if isempty(startTimes)
    return;
end
yLimits = ylim(axisHandle);
for iInt = 1:numel(startTimes)
    currentPatch = patch(axisHandle, [startTimes(iInt) endTimes(iInt) endTimes(iInt) startTimes(iInt)], ...
        [yLimits(1) yLimits(1) yLimits(2) yLimits(2)], colorValue, ...
        'FaceAlpha', alphaValue, 'EdgeColor', 'none');
    if iInt == 1
        patchHandle = currentPatch;
    end
end
uistack(patchHandle, 'bottom');
end

function entries = localSpeedLegend(hasReference, hasActual)
entries = {};
if hasReference
    entries{end + 1} = 'v_{ref}';
end
if hasActual
    entries{end + 1} = 'v_{act}';
end
end

function entries = localPedalLegend(hasAccel, hasBrake)
entries = {};
if hasAccel
    entries{end + 1} = 'Accel';
end
if hasBrake
    entries{end + 1} = 'Brake';
end
end

function [handles, labels] = localEventLegend(refHandle, actHandle, accPatch, brkPatch)
handles = [];
labels = {};
if isgraphics(refHandle)
    handles(end + 1) = refHandle; %#ok<AGROW>
    labels{end + 1} = 'v_{ref}'; %#ok<AGROW>
end
if isgraphics(actHandle)
    handles(end + 1) = actHandle; %#ok<AGROW>
    labels{end + 1} = 'v_{act}'; %#ok<AGROW>
end
if isgraphics(accPatch)
    handles(end + 1) = accPatch; %#ok<AGROW>
    labels{end + 1} = 'Accel events'; %#ok<AGROW>
end
if isgraphics(brkPatch)
    handles(end + 1) = brkPatch; %#ok<AGROW>
    labels{end + 1} = 'Brake events'; %#ok<AGROW>
end
end

function value = localDominantValue(data)
data = data(isfinite(data));
if isempty(data)
    value = NaN;
    return;
end
uniqueValues = unique(data);
counts = zeros(size(uniqueValues));
for iValue = 1:numel(uniqueValues)
    counts(iValue) = sum(data == uniqueValues(iValue));
end
[~, idx] = max(counts);
value = uniqueValues(idx);
end

function gradeClass = localGradeClass(meanSlope, config)
if ~isfinite(meanSlope)
    gradeClass = "Unknown";
elseif meanSlope > config.Thresholds.SteepSlope_pct
    gradeClass = "Steep Uphill";
elseif meanSlope > config.Thresholds.UphillSlope_pct
    gradeClass = "Uphill";
elseif meanSlope < -config.Thresholds.SteepSlope_pct
    gradeClass = "Steep Downhill";
elseif meanSlope < config.Thresholds.DownhillSlope_pct
    gradeClass = "Downhill";
else
    gradeClass = "Flat";
end
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
