function [vehicleKPI, narrative] = RCA_ComputeVehicleKPIs(derived, ~, ~, signalPresence, config)
% RCA_ComputeVehicleKPIs  Compute trip-level vehicle KPIs.

t = derived.time_s;
rows = cell(0, 7);
narrative = strings(0, 1);

tripDuration = max(t) - min(t);
tripDistance = derived.tripDistance_km;
validSpeed = isfinite(derived.vehVel_kmh);
stopShare = 100 * RCA_FractionTrue(derived.vehVel_kmh <= config.Thresholds.StopSpeed_kmh, validSpeed);
rows = RCA_AddKPI(rows, 'Trip Duration', tripDuration, 's', 'DriveCycle', 'Vehicle', 'reference time', 'Full trip duration.');
rows = RCA_AddKPI(rows, 'Trip Distance', tripDistance, 'km', 'DriveCycle', 'Vehicle', 'veh_pos or integrated veh_vel', 'Distance basis selected automatically.');
rows = RCA_AddKPI(rows, 'Average Speed', mean(derived.vehVel_kmh, 'omitnan'), 'km/h', 'DriveCycle', 'Vehicle', 'veh_vel', 'Trip mean vehicle speed.');
rows = RCA_AddKPI(rows, 'Peak Speed', max(derived.vehVel_kmh, [], 'omitnan'), 'km/h', 'DriveCycle', 'Vehicle', 'veh_vel', 'Trip peak vehicle speed.');
rows = RCA_AddKPI(rows, 'Stop Time Share', stopShare, '%', 'DriveCycle', 'Vehicle', 'veh_vel', 'Stop threshold is defined in RCA_Config.');

if isfield(derived, 'targetDistance_km') && isfinite(derived.targetDistance_km) && derived.targetDistance_km > 0
    targetDistance = derived.targetDistance_km;
    distanceError = tripDistance - targetDistance;
    distanceErrorAbs = abs(distanceError);
    distanceErrorPct = 100 * distanceErrorAbs / max(abs(targetDistance), eps);
    sourceText = 'cfg_target_distance + trip distance';
    if isfield(derived, 'targetDistanceSource') && strlength(string(derived.targetDistanceSource)) > 0
        sourceText = sprintf('%s + trip distance', char(string(derived.targetDistanceSource)));
    end
    rows = RCA_AddKPI(rows, 'Target Distance', targetDistance, 'km', 'Performance', 'Vehicle', sourceText, 'Target distance from run configuration; unit normalized to km when required.');
    rows = RCA_AddKPI(rows, 'Target-Actual Distance Error', distanceError, 'km', 'Performance', 'Vehicle', sourceText, 'Actual trip distance minus configured target distance.');
    rows = RCA_AddKPI(rows, 'Absolute Distance Error', distanceErrorAbs, 'km', 'Performance', 'Vehicle', sourceText, 'Absolute difference between configured target distance and achieved trip distance.');
    rows = RCA_AddKPI(rows, 'Distance Error Percentage', distanceErrorPct, '%', 'Performance', 'Vehicle', sourceText, 'Absolute distance error normalized by configured target distance.');
end

if ~all(isnan(derived.speedDemand_kmh))
    speedErr = derived.speedDemand_kmh - derived.vehVel_kmh;
    validSpeedErr = isfinite(speedErr);
    rows = RCA_AddKPI(rows, 'Speed Tracking MAE', mean(abs(speedErr), 'omitnan'), 'km/h', 'Performance', 'Vehicle', 'veh_des_vel + veh_vel', 'Mean absolute tracking error.');
    rows = RCA_AddKPI(rows, 'Speed Tracking RMSE', sqrt(mean(speedErr .^ 2, 'omitnan')), 'km/h', 'Performance', 'Vehicle', 'veh_des_vel + veh_vel', 'Root mean square tracking error.');
    rows = RCA_AddKPI(rows, 'Time Above Tracking Threshold', ...
        100 * RCA_FractionTrue(abs(speedErr) > config.Thresholds.PoorTrackingError_kmh, validSpeedErr), '%', ...
        'Performance', 'Vehicle', 'veh_des_vel + veh_vel', 'Tracking threshold is configured in RCA_Config.');
end

dischargeEnergy = RCA_TrapzFinite(t, max(derived.batteryPower_kW, 0)) / 3600;
regenEnergy = RCA_TrapzFinite(t, max(-derived.batteryPower_kW, 0)) / 3600;
netEnergy = dischargeEnergy - regenEnergy;
rows = RCA_AddKPI(rows, 'Battery Discharge Energy', dischargeEnergy, 'kWh', 'Energy', 'Vehicle', 'batt_pwr', 'Integrated discharge-positive battery power after applying workbook sign convention.');
rows = RCA_AddKPI(rows, 'Battery Regen Energy', regenEnergy, 'kWh', 'Energy', 'Vehicle', 'batt_pwr', 'Integrated charging/recovered battery power after applying workbook sign convention.');
rows = RCA_AddKPI(rows, 'Net Battery Energy', netEnergy, 'kWh', 'Energy', 'Vehicle', 'batt_pwr', 'Discharge minus recovered electrical energy using workbook sign convention.');

if tripDistance > config.General.MinimumDistanceForWhpkm_km
    rows = RCA_AddKPI(rows, 'Energy Intensity', netEnergy * 1000 / tripDistance, 'Wh/km', 'Efficiency', 'Vehicle', 'batt_pwr + trip distance', 'Net electrical energy per kilometre.');
end

auxEnergy = RCA_TrapzFinite(t, max(derived.auxiliaryPower_kW, 0)) / 3600;
hprEnergy = RCA_TrapzFinite(t, max(derived.highPowerResistorPower_kW, 0)) / 3600;
tractionEnergy = RCA_TrapzFinite(t, max(derived.tractionPower_kW, 0)) / 3600;
motorLossEnergy = RCA_TrapzFinite(t, max(derived.motorLossPower_kW, 0)) / 3600;
gbxLossEnergy = RCA_TrapzFinite(t, max(derived.gearboxLossPower_kW, 0)) / 3600;
battLossEnergy = RCA_TrapzFinite(t, max(derived.batteryLossPower_kW, 0)) / 3600;
fricEnergy = RCA_TrapzFinite(t, max(derived.frictionBrakePower_kW, 0)) / 3600;

rows = RCA_AddKPI(rows, 'Auxiliary Energy', auxEnergy, 'kWh', 'Energy', 'Vehicle', 'aux_curr + aux_volt', 'Integrated auxiliary electrical demand.');
rows = RCA_AddKPI(rows, 'High Power Resistor Energy', hprEnergy, 'kWh', 'Energy', 'Vehicle', 'hpr_pwr', 'Integrated high-power resistor energy when available.');
rows = RCA_AddKPI(rows, 'Traction Mechanical Energy', tractionEnergy, 'kWh', 'Energy', 'Vehicle', 'traction force + vehicle speed', 'Wheel-end mechanical work.');
rows = RCA_AddKPI(rows, 'Battery Loss Energy', battLossEnergy, 'kWh', 'Losses', 'Vehicle', 'batt_loss_pwr', 'Battery internal/system loss integral.');
rows = RCA_AddKPI(rows, 'Motor/Inverter Loss Energy', motorLossEnergy, 'kWh', 'Losses', 'Vehicle', 'emot loss power', 'Electric drive loss integral.');
rows = RCA_AddKPI(rows, 'Transmission Loss Energy', gbxLossEnergy, 'kWh', 'Losses', 'Vehicle', 'gbx_pwr_loss', 'Gearbox loss integral.');
rows = RCA_AddKPI(rows, 'Friction Brake Energy', fricEnergy, 'kWh', 'Losses', 'Vehicle', 'fric_brk_pwr', 'Friction dissipation integral.');
if dischargeEnergy > 0
    batteryToWheelEff = 100 * tractionEnergy / dischargeEnergy;
    auxiliaryShare = 100 * auxEnergy / dischargeEnergy;
else
    batteryToWheelEff = NaN;
    auxiliaryShare = NaN;
end
if (regenEnergy + fricEnergy) > 0
    regenRecovery = 100 * regenEnergy / (regenEnergy + fricEnergy);
else
    regenRecovery = NaN;
end
rows = RCA_AddKPI(rows, 'Battery-to-Wheel Efficiency', batteryToWheelEff, '%', 'Efficiency', 'Vehicle', 'batt_pwr + traction power', 'Mechanical wheel output over discharge-positive battery energy.');
rows = RCA_AddKPI(rows, 'Auxiliary Energy Share', auxiliaryShare, '%', 'Efficiency', 'Vehicle', 'auxiliary power + battery power', 'Auxiliary share of discharge-positive battery energy.');
rows = RCA_AddKPI(rows, 'Approximate Regen Recovery Fraction', regenRecovery, '%', 'Efficiency', 'Vehicle', 'batt_pwr + fric_brk_pwr', 'Recovered electrical braking divided by recovered plus friction braking energy after sign normalization.');

terminalPowerResidual = derived.powerBalanceResidualTerminal_kW(:);
internalPowerResidual = derived.powerBalanceResidualInternal_kW(:);
wheelForceResidual = derived.forceBalanceResidualWheel_N(:);
roadLoadForceResidual = derived.forceBalanceResidualRoadLoad_N(:);
validPowerResidual = isfinite(terminalPowerResidual);
validForceResidual = isfinite(wheelForceResidual);

rows = RCA_AddKPI(rows, 'Power Balance Residual MAE (Terminal)', mean(abs(terminalPowerResidual), 'omitnan'), 'kW', ...
    'Balance', 'Vehicle', 'battery power - (motor electrical + auxiliary + HPR)', ...
    'Residual of terminal-level electrical power balance. Lower values indicate better agreement between source and sink powers.');
rows = RCA_AddKPI(rows, 'Power Balance Residual 95th Percentile (Terminal)', RCA_Percentile(abs(terminalPowerResidual(validPowerResidual)), 95), 'kW', ...
    'Balance', 'Vehicle', 'battery power - (motor electrical + auxiliary + HPR)', ...
    'Tail severity of terminal-level electrical balance mismatch.');
rows = RCA_AddKPI(rows, 'Power Balance Residual MAE (Internal)', mean(abs(internalPowerResidual), 'omitnan'), 'kW', ...
    'Balance', 'Vehicle', 'battery power - battery loss - (motor electrical + auxiliary + HPR)', ...
    'Residual after including battery loss in the electrical power balance.');
rows = RCA_AddKPI(rows, 'Force Balance Residual MAE (Wheel Path)', mean(abs(wheelForceResidual), 'omitnan'), 'N', ...
    'Balance', 'Vehicle', 'wheel force - (vehicle propulsion force - friction brake force)', ...
    'Residual between wheel force and propulsion-minus-braking force path.');
rows = RCA_AddKPI(rows, 'Force Balance Residual 95th Percentile (Wheel Path)', RCA_Percentile(abs(wheelForceResidual(validForceResidual)), 95), 'N', ...
    'Balance', 'Vehicle', 'wheel force - (vehicle propulsion force - friction brake force)', ...
    'Tail severity of the wheel-path force-balance mismatch.');
rows = RCA_AddKPI(rows, 'Force Balance Residual MAE (Road Load)', mean(abs(roadLoadForceResidual), 'omitnan'), 'N', ...
    'Balance', 'Vehicle', 'wheel force - (rolling + grade + aero + inertial force)', ...
    'Residual between wheel force and summed road-load plus inertial forces.');

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
    finiteSoc = derived.batterySOC_pct(isfinite(derived.batterySOC_pct));
    usableSoc = NaN;
    remainingSoc = NaN;
    if numel(finiteSoc) >= 2
        usableSoc = max(finiteSoc(1) - finiteSoc(end), 0);
        remainingSoc = max(finiteSoc(end) - config.General.DefaultReserveSOC_pct, 0);
    end
    if usableSoc > 0 && tripDistance > 0
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
narrative(end + 1) = sprintf('Balance checks: terminal power-balance MAE %.1f kW, wheel-path force-balance MAE %.0f N, and road-load force-balance MAE %.0f N.', ...
    mean(abs(terminalPowerResidual), 'omitnan'), mean(abs(wheelForceResidual), 'omitnan'), mean(abs(roadLoadForceResidual), 'omitnan'));
narrative(end + 1) = sprintf('Gear behaviour: %d shifts, %.1f shifts/km, %d detected hunting events.', ...
    shiftCount, shiftCount / max(tripDistance, eps), huntingCount);

vehicleKPI = RCA_FinalizeKPITable(rows);
end
