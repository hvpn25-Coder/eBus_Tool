function result = Analyze_Transmission(analysisData, outputPaths, config)
% Analyze_Transmission  Gear usage, shift quality, and gearbox loss analysis.

result = localInitResult("TRANSMISSION", {'gr_num'}, {'gr_ratio', 'gbx_pwr_loss', 'emot1_act_spd', 'emot2_act_spd', 'veh_vel'});
t = analysisData.Derived.time_s;
gear = analysisData.Derived.gearNumber;
vehSpeed = analysisData.Derived.vehVel_kmh;
motorSpd = analysisData.Derived.motorSpeed_rpm;
gbxLoss = analysisData.Derived.gearboxLossPower_kW;
motorElec = analysisData.Derived.motorElectricalPower_kW;

rows = cell(0, 7);
summary = strings(0, 1);
if all(isnan(gear))
    result.Warnings(end + 1) = "Actual gear signal is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Transmission", strings(0, 1), strings(0, 1));
    return;
end

changeIdx = find(abs(diff(gear)) > 0 & ~isnan(diff(gear))) + 1;
shiftCount = numel(changeIdx);
distanceKm = max(analysisData.Derived.tripDistance_km, eps);
shiftRate = shiftCount / distanceKm;
dwell = diff([t(1); t(changeIdx); t(end)]);
huntingCount = 0;
for iShift = 3:numel(changeIdx)
    if gear(changeIdx(iShift)) == gear(changeIdx(iShift - 2)) && ...
            (t(changeIdx(iShift)) - t(changeIdx(iShift - 2))) <= config.Thresholds.GearHuntingWindow_s
        huntingCount = huntingCount + 1;
    end
end

rows = RCA_AddKPI(rows, 'Gear Shift Count', shiftCount, 'count', 'Operation', 'Transmission', 'gr_num', 'Shift count from gear state transitions.');
rows = RCA_AddKPI(rows, 'Gear Shift Rate', shiftRate, 'shifts/km', 'Operation', 'Transmission', 'gr_num + trip distance', 'High shift density suggests unstable ratio usage.');
rows = RCA_AddKPI(rows, 'Mean Gear Dwell', mean(dwell, 'omitnan'), 's', 'Operation', 'Transmission', 'gr_num', 'Low dwell indicates unstable gear holding.');
rows = RCA_AddKPI(rows, 'Gear Hunting Count', huntingCount, 'count', 'Operation', 'Transmission', 'gr_num', 'A-B-A gear reversals within the configured hunting window.');
rows = RCA_AddKPI(rows, 'Average Gearbox Loss Power', mean(max(gbxLoss, 0), 'omitnan'), 'kW', 'Losses', 'Transmission', 'gbx_pwr_loss', 'Available when gearbox loss signal is logged.');

effByGearRows = strings(0, 1);
uniqueGears = unique(gear(~isnan(gear)));
energyByGear = NaN(size(uniqueGears));
for iGear = 1:numel(uniqueGears)
    mask = gear == uniqueGears(iGear);
    energyByGear(iGear) = trapz(t(mask), max(motorElec(mask), 0)) / 3600;
    effByGearRows(end + 1) = sprintf('Gear %.0f drive-positive electrical energy usage is %.2f kWh.', uniqueGears(iGear), energyByGear(iGear));
end
summary = [summary; effByGearRows];

recs = strings(0, 1);
evidence = strings(0, 1);
if shiftRate > config.Thresholds.GearShiftRate_perkm
    recs(end + 1) = "Reduce shift activity by retuning the shift schedule or hysteresis; current shift density is high for efficiency and drivability.";
    evidence(end + 1) = sprintf('Shift rate is %.1f shifts/km against a heuristic threshold of %.1f.', shiftRate, config.Thresholds.GearShiftRate_perkm);
end
if mean(dwell, 'omitnan') < config.Thresholds.MinGearDwell_s
    recs(end + 1) = "Increase gear dwell or add hysteresis to reduce unstable gear toggling.";
    evidence(end + 1) = sprintf('Mean dwell is %.1f s, below the %.1f s heuristic.', mean(dwell, 'omitnan'), config.Thresholds.MinGearDwell_s);
end
if huntingCount > 0
    recs(end + 1) = "Investigate gear hunting around load transitions; repeated ratio reversals waste energy and disrupt torque delivery.";
    evidence(end + 1) = sprintf('%d gear-hunting events were detected.', huntingCount);
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 2, 1);
stairs(t, gear, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
title('Gear Number Versus Time');
xlabel('Time (s)');
ylabel('Gear');
grid on;

subplot(2, 2, 2);
scatter(vehSpeed, gear, 12, motorSpd, 'filled');
title('Gear Versus Vehicle Speed');
xlabel('Vehicle Speed (km/h)');
ylabel('Gear');
cb = colorbar;
cb.Label.String = 'Motor speed (rpm)';
grid on;

subplot(2, 2, 3);
histogram(gear(~isnan(gear)), 'FaceColor', config.Plot.Colors.Gear);
title('Gear Usage Histogram');
xlabel('Gear');
ylabel('Samples');
grid on;

subplot(2, 2, 4);
bar(uniqueGears, energyByGear, 'FaceColor', config.Plot.Colors.Vehicle);
title('Electrical Energy by Gear');
xlabel('Gear');
ylabel('Energy (kWh)');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Transmission'), 'Transmission_Gear_Analysis', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Transmission", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
