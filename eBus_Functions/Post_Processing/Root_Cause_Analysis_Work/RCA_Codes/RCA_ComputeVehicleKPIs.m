function [vehicleKPI, narrative] = RCA_ComputeVehicleKPIs(derived, ~, ~, signalPresence, config)
% RCA_ComputeVehicleKPIs  Compute trip-level vehicle KPIs.

t = derived.time_s;
rows = cell(0, 7);
narrative = strings(0, 1);

tripDuration = max(t) - min(t);
tripDistance = derived.tripDistance_km;
stopShare = 100 * mean(derived.vehVel_kmh <= config.Thresholds.StopSpeed_kmh, 'omitnan');
rows = RCA_AddKPI(rows, 'Trip Duration', tripDuration, 's', 'DriveCycle', 'Vehicle', 'reference time', 'Full trip duration.');
rows = RCA_AddKPI(rows, 'Trip Distance', tripDistance, 'km', 'DriveCycle', 'Vehicle', 'veh_pos or integrated veh_vel', 'Distance basis selected automatically.');
rows = RCA_AddKPI(rows, 'Average Speed', mean(derived.vehVel_kmh, 'omitnan'), 'km/h', 'DriveCycle', 'Vehicle', 'veh_vel', 'Trip mean vehicle speed.');
rows = RCA_AddKPI(rows, 'Peak Speed', max(derived.vehVel_kmh, [], 'omitnan'), 'km/h', 'DriveCycle', 'Vehicle', 'veh_vel', 'Trip peak vehicle speed.');
rows = RCA_AddKPI(rows, 'Stop Time Share', stopShare, '%', 'DriveCycle', 'Vehicle', 'veh_vel', 'Stop threshold is defined in RCA_Config.');

if ~all(isnan(derived.speedDemand_kmh))
    speedErr = derived.speedDemand_kmh - derived.vehVel_kmh;
    rows = RCA_AddKPI(rows, 'Speed Tracking MAE', mean(abs(speedErr), 'omitnan'), 'km/h', 'Performance', 'Vehicle', 'veh_des_vel + veh_vel', 'Mean absolute tracking error.');
    rows = RCA_AddKPI(rows, 'Speed Tracking RMSE', sqrt(mean(speedErr .^ 2, 'omitnan')), 'km/h', 'Performance', 'Vehicle', 'veh_des_vel + veh_vel', 'Root mean square tracking error.');
    rows = RCA_AddKPI(rows, 'Time Above Tracking Threshold', ...
        100 * mean(abs(speedErr) > config.Thresholds.PoorTrackingError_kmh, 'omitnan'), '%', ...
        'Performance', 'Vehicle', 'veh_des_vel + veh_vel', 'Tracking threshold is configured in RCA_Config.');
end

dischargeEnergy = trapz(t, max(derived.batteryPower_kW, 0)) / 3600;
regenEnergy = trapz(t, max(-derived.batteryPower_kW, 0)) / 3600;
netEnergy = dischargeEnergy - regenEnergy;
rows = RCA_AddKPI(rows, 'Battery Discharge Energy', dischargeEnergy, 'kWh', 'Energy', 'Vehicle', 'batt_pwr', 'Integrated positive battery power.');
rows = RCA_AddKPI(rows, 'Battery Regen Energy', regenEnergy, 'kWh', 'Energy', 'Vehicle', 'batt_pwr', 'Integrated recovered battery power.');
rows = RCA_AddKPI(rows, 'Net Battery Energy', netEnergy, 'kWh', 'Energy', 'Vehicle', 'batt_pwr', 'Discharge minus recovered electrical energy.');

if tripDistance > config.General.MinimumDistanceForWhpkm_km
    rows = RCA_AddKPI(rows, 'Energy Intensity', netEnergy * 1000 / tripDistance, 'Wh/km', 'Efficiency', 'Vehicle', 'batt_pwr + trip distance', 'Net electrical energy per kilometre.');
end

auxEnergy = trapz(t, max(derived.auxiliaryPower_kW, 0)) / 3600;
tractionEnergy = trapz(t, max(derived.tractionPower_kW, 0)) / 3600;
motorLossEnergy = trapz(t, max(derived.motorLossPower_kW, 0)) / 3600;
gbxLossEnergy = trapz(t, max(derived.gearboxLossPower_kW, 0)) / 3600;
battLossEnergy = trapz(t, max(derived.batteryLossPower_kW, 0)) / 3600;
fricEnergy = trapz(t, max(derived.frictionBrakePower_kW, 0)) / 3600;

rows = RCA_AddKPI(rows, 'Auxiliary Energy', auxEnergy, 'kWh', 'Energy', 'Vehicle', 'aux_curr + aux_volt', 'Integrated auxiliary electrical demand.');
rows = RCA_AddKPI(rows, 'Traction Mechanical Energy', tractionEnergy, 'kWh', 'Energy', 'Vehicle', 'traction force + vehicle speed', 'Wheel-end mechanical work.');
rows = RCA_AddKPI(rows, 'Battery Loss Energy', battLossEnergy, 'kWh', 'Losses', 'Vehicle', 'batt_loss_pwr', 'Battery internal/system loss integral.');
rows = RCA_AddKPI(rows, 'Motor/Inverter Loss Energy', motorLossEnergy, 'kWh', 'Losses', 'Vehicle', 'emot loss power', 'Electric drive loss integral.');
rows = RCA_AddKPI(rows, 'Transmission Loss Energy', gbxLossEnergy, 'kWh', 'Losses', 'Vehicle', 'gbx_pwr_loss', 'Gearbox loss integral.');
rows = RCA_AddKPI(rows, 'Friction Brake Energy', fricEnergy, 'kWh', 'Losses', 'Vehicle', 'fric_brk_pwr', 'Friction dissipation integral.');
rows = RCA_AddKPI(rows, 'Battery-to-Wheel Efficiency', 100 * tractionEnergy / max(dischargeEnergy, eps), '%', 'Efficiency', 'Vehicle', 'batt_pwr + traction power', 'Mechanical wheel output over electrical battery discharge.');
rows = RCA_AddKPI(rows, 'Auxiliary Energy Share', 100 * auxEnergy / max(dischargeEnergy, eps), '%', 'Efficiency', 'Vehicle', 'auxiliary power + battery power', 'Auxiliary share of positive battery discharge energy.');
rows = RCA_AddKPI(rows, 'Approximate Regen Recovery Fraction', 100 * regenEnergy / max(regenEnergy + fricEnergy, eps), '%', 'Efficiency', 'Vehicle', 'batt_pwr + fric_brk_pwr', 'Recovered electrical braking divided by recovered plus friction braking energy.');

gear = derived.gearNumber;
changeIdx = find(abs(diff(gear)) > 0 & ~isnan(diff(gear))) + 1;
shiftCount = numel(changeIdx);
huntingCount = 0;
for iShift = 3:numel(changeIdx)
    if gear(changeIdx(iShift)) == gear(changeIdx(iShift - 2)) && ...
            (t(changeIdx(iShift)) - t(changeIdx(iShift - 2))) <= config.Thresholds.GearHuntingWindow_s
        huntingCount = huntingCount + 1;
    end
end
rows = RCA_AddKPI(rows, 'Gear Shift Count', shiftCount, 'count', 'Gear', 'Vehicle', 'gr_num', 'Gear transitions counted from the logged gear signal.');
rows = RCA_AddKPI(rows, 'Gear Shift Rate', shiftCount / max(tripDistance, eps), 'shifts/km', 'Gear', 'Vehicle', 'gr_num + trip distance', 'High shift density can signal poor ratio usage.');
rows = RCA_AddKPI(rows, 'Gear Hunting Count', huntingCount, 'count', 'Gear', 'Vehicle', 'gr_num', 'A-B-A reversals within the configured hunting window.');

if ~all(isnan(derived.batterySOC_pct))
    usableSoc = max(derived.batterySOC_pct(1) - derived.batterySOC_pct(end), 0);
    if usableSoc > 0 && tripDistance > 0
        remainingSoc = max(derived.batterySOC_pct(end) - config.General.DefaultReserveSOC_pct, 0);
        estimatedRemainingRange = tripDistance / usableSoc * remainingSoc;
        rows = RCA_AddKPI(rows, 'Estimated Remaining Range', estimatedRemainingRange, 'km', 'Range', 'Vehicle', 'batt_soc + trip distance', ...
            'Empirical estimate based on observed distance per percentage-point SoC.');
    end
end

rows = RCA_AddKPI(rows, 'Present Signal Count', sum(signalPresence.Status == "Present"), 'count', 'SignalCoverage', 'Vehicle', 'signal presence table', 'Signals resolved from workbook or fallback search.');
rows = RCA_AddKPI(rows, 'Missing Required Signal Count', sum(signalPresence.Status == "Missing"), 'count', 'SignalCoverage', 'Vehicle', 'signal presence table', 'Missing required signals reduce RCA confidence.');
rows = RCA_AddKPI(rows, 'Optional Missing Signal Count', sum(signalPresence.Status == "Optional Missing"), 'count', 'SignalCoverage', 'Vehicle', 'signal presence table', 'Optional missing signals do not stop execution but limit detail.');

narrative(end + 1) = sprintf('Trip overview: %.2f km over %.1f s with %.1f km/h average speed and %.2f kWh net battery energy.', ...
    tripDistance, tripDuration, mean(derived.vehVel_kmh, 'omitnan'), netEnergy);
narrative(end + 1) = sprintf('Energy split: traction %.2f kWh, auxiliaries %.2f kWh, battery loss %.2f kWh, motor loss %.2f kWh, gearbox loss %.2f kWh.', ...
    tractionEnergy, auxEnergy, battLossEnergy, motorLossEnergy, gbxLossEnergy);
narrative(end + 1) = sprintf('Gear behaviour: %d shifts, %.1f shifts/km, %d detected hunting events.', ...
    shiftCount, shiftCount / max(tripDistance, eps), huntingCount);

vehicleKPI = RCA_FinalizeKPITable(rows);
end
