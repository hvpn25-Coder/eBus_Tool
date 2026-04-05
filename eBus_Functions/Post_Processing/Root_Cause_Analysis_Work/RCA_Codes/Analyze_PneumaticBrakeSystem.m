function result = Analyze_PneumaticBrakeSystem(analysisData, outputPaths, config)
% Analyze_PneumaticBrakeSystem  Friction-brake usage, braking severity, and regen substitution RCA.

result = localInitResult("PNEUMATIC BRAKE SYSTEM", ...
    {'fric_brk_force', 'fric_brk_pwr'}, ...
    {'brk_pdl', 'batt_pwr', 'veh_vel', 'road_slp'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Pneumatic brake analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Pneumatic Brake System", recs, evidence);
    result.SummaryText = summary;
    return;
end

fricForce = d.frictionBrakeForce_N(:);
fricPower = d.frictionBrakePower_kW(:);
vehSpeed = d.vehVel_kmh(:);
brkPedal = d.brkPedal_pct(:);
roadSlope = d.roadSlope_pct(:);
regenElec = d.motorRegenElectricalPowerPositive_kW(:);
regenMech = d.motorRegenMechanicalPowerPositive_kW(:);
battCharge = d.batteryChargePowerPositive_kW(:);

if ~any(isfinite(fricForce))
    fricForce = localAlignedSignal(analysisData.Signals, 'fric_brk_force', n);
end
if ~any(isfinite(fricPower))
    fricPower = localAlignedSignal(analysisData.Signals, 'fric_brk_pwr', n);
end

if all(isnan(fricForce)) && all(isnan(fricPower))
    result.Warnings(end + 1) = "Pneumatic brake analysis skipped because friction brake force and power are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Pneumatic Brake System", recs, evidence);
    result.SummaryText = summary;
    return;
end

movingMask = isfinite(vehSpeed) & vehSpeed > config.Thresholds.StopSpeed_kmh;
fricActiveMask = isfinite(fricPower) & fricPower > 1;
fricForceActiveMask = isfinite(fricForce) & abs(fricForce) > 50;
brakePedalMask = isfinite(brkPedal) & brkPedal >= config.Thresholds.DriverPedalActive_pct;
regenOpportunityMask = movingMask & isfinite(regenMech) & regenMech >= config.Thresholds.RegenOpportunityBrakePower_kW;
highFricPowerMask = movingMask & isfinite(fricPower) & fricPower >= config.Thresholds.RegenOpportunityBrakePower_kW;
downhillBrakeMask = brakePedalMask & isfinite(roadSlope) & roadSlope <= config.Thresholds.DownhillSlope_pct;

fricEnergy = RCA_TrapzFinite(t, max(fricPower, 0)) / 3600;
regenElectricalEnergy = RCA_TrapzFinite(t, max(regenElec, 0)) / 3600;
battChargeEnergy = RCA_TrapzFinite(t, max(battCharge, 0)) / 3600;
regenMechanicalEnergy = RCA_TrapzFinite(t, max(regenMech, 0)) / 3600;
brakeOpportunity = fricEnergy + regenElectricalEnergy;
regenRecovery = NaN;
if brakeOpportunity > 0
    regenRecovery = 100 * regenElectricalEnergy / brakeOpportunity;
end

fricShareOfBraking = NaN(size(fricPower));
validBrakingMix = isfinite(fricPower) & isfinite(regenElec) & (max(fricPower, 0) + max(regenElec, 0)) > 1;
fricShareOfBraking(validBrakingMix) = 100 * max(fricPower(validBrakingMix), 0) ./ ...
    max(max(fricPower(validBrakingMix), 0) + max(regenElec(validBrakingMix), 0), eps);

rows = RCA_AddKPI(rows, 'Friction Brake Energy', fricEnergy, 'kWh', ...
    'Losses', 'Pneumatic Brake System', 'fric_brk_pwr', ...
    'Integrated positive friction brake power.');
rows = RCA_AddKPI(rows, 'Peak Friction Brake Force', max(abs(fricForce), [], 'omitnan'), 'N', ...
    'Performance', 'Pneumatic Brake System', 'fric_brk_force', ...
    'Maximum magnitude of friction brake force.');
rows = RCA_AddKPI(rows, 'Peak Friction Brake Power', max(fricPower, [], 'omitnan'), 'kW', ...
    'Performance', 'Pneumatic Brake System', 'fric_brk_pwr', ...
    'Maximum positive friction brake power.');
rows = RCA_AddKPI(rows, 'Friction Brake Active Share', 100 * RCA_FractionTrue(fricActiveMask, movingMask), '%', ...
    'Operation', 'Pneumatic Brake System', 'fric_brk_pwr + veh_vel', ...
    'Share of moving samples where friction brake power is active.');
rows = RCA_AddKPI(rows, 'Brake Pedal With Friction Share', 100 * RCA_FractionTrue(fricActiveMask, brakePedalMask), '%', ...
    'Operation', 'Pneumatic Brake System', 'brk_pdl + fric_brk_pwr', ...
    'Share of brake-pedal-active samples where friction braking is engaged.');
rows = RCA_AddKPI(rows, 'Approximate Regen Recovery Fraction', regenRecovery, '%', ...
    'Efficiency', 'Pneumatic Brake System', 'fric_brk_pwr + motor regen electrical power', ...
    'Electrical recovery divided by friction plus electrical braking opportunity.');
rows = RCA_AddKPI(rows, 'Average Friction Share of Braking', mean(fricShareOfBraking, 'omitnan'), '%', ...
    'Efficiency', 'Pneumatic Brake System', 'fric_brk_pwr + motor regen electrical power', ...
    'Average share of braking power handled by friction when braking power is present.');
rows = RCA_AddKPI(rows, 'Brake Power Above Regen Opportunity Threshold', ...
    100 * RCA_FractionTrue(highFricPowerMask, movingMask), '%', ...
    'Efficiency', 'Pneumatic Brake System', 'fric_brk_pwr + veh_vel', ...
    sprintf('Share of moving samples where friction brake power exceeds %.1f kW.', config.Thresholds.RegenOpportunityBrakePower_kW));
rows = RCA_AddKPI(rows, 'Downhill Friction Brake Share', 100 * RCA_FractionTrue(fricActiveMask, downhillBrakeMask), '%', ...
    'Operation', 'Pneumatic Brake System', 'fric_brk_pwr + road_slope', ...
    'Share of downhill brake-pedal-active samples using friction braking.');
rows = RCA_AddKPI(rows, 'Regenerative Mechanical Braking Energy', regenMechanicalEnergy, 'kWh', ...
    'Efficiency', 'Pneumatic Brake System', 'motor mechanical regen power', ...
    'Mechanical braking energy absorbed by the electric machines during regeneration.');
rows = RCA_AddKPI(rows, 'Battery Charge Energy During Braking', battChargeEnergy, 'kWh', ...
    'Efficiency', 'Pneumatic Brake System', 'battery charge power', ...
    'Battery charge energy accumulated during braking-capable operation.');
rows = RCA_AddKPI(rows, 'Mean Friction Brake Force While Active', mean(abs(fricForce(fricForceActiveMask)), 'omitnan'), 'N', ...
    'Operation', 'Pneumatic Brake System', 'fric_brk_force', ...
    'Average magnitude of friction brake force while active.');

summary(end + 1) = sprintf(['Pneumatic braking dissipates %.2f kWh as friction heat with peak friction brake power of %.1f kW. ', ...
    'Friction braking is active for %.1f%% of moving samples.'], ...
    fricEnergy, max(fricPower, [], 'omitnan'), 100 * RCA_FractionTrue(fricActiveMask, movingMask));
summary(end + 1) = sprintf(['Braking energy split: approximate electrical recovery is %.2f kWh and estimated regen recovery fraction is %.1f%%. ', ...
    'Average friction share of braking is %.1f%%.'], ...
    regenElectricalEnergy, regenRecovery, mean(fricShareOfBraking, 'omitnan'));
summary(end + 1) = sprintf(['Braking context: downhill friction brake share is %.1f%% and high friction-brake power share is %.1f%% above %.1f kW.'], ...
    100 * RCA_FractionTrue(fricActiveMask, downhillBrakeMask), ...
    100 * RCA_FractionTrue(highFricPowerMask, movingMask), config.Thresholds.RegenOpportunityBrakePower_kW);

if regenRecovery < config.Thresholds.PoorRegenRecoveryFraction * 100
    recs(end + 1) = "Improve brake blending or regen availability; friction braking is consuming too much of the observable braking energy opportunity.";
    evidence(end + 1) = sprintf('Approximate regen recovery is %.1f%%.', regenRecovery);
end
if mean(fricShareOfBraking, 'omitnan') > 60
    recs(end + 1) = "Review regen substitution strategy because friction braking carries a large average share of braking duty.";
    evidence(end + 1) = sprintf('Average friction share of braking is %.1f%%.', mean(fricShareOfBraking, 'omitnan'));
end
if 100 * RCA_FractionTrue(fricActiveMask, downhillBrakeMask) > 20
    recs(end + 1) = "Investigate downhill braking calibration and charge-acceptance limits. Friction braking is frequently used on downhill brake events where regen opportunity should be high.";
    evidence(end + 1) = sprintf('Downhill friction brake share is %.1f%%.', 100 * RCA_FractionTrue(fricActiveMask, downhillBrakeMask));
end
if fricEnergy > 0.5
    recs(end + 1) = "Inspect braking-heavy segments for excess friction dissipation and identify whether the limitation is battery charge acceptance, motor negative torque limit, or control blending.";
    evidence(end + 1) = sprintf('Integrated friction brake energy is %.2f kWh.', fricEnergy);
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'PneumaticBrakeSystem');
plotFiles = localAppendPlotFile(plotFiles, localPlotOverview(figureFolder, t, fricForce, fricPower, brkPedal, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotBrakeBlending(figureFolder, t, fricPower, regenElec, battCharge, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotBrakeContext(figureFolder, vehSpeed, roadSlope, fricForce, fricPower, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotEnergyBreakdown(figureFolder, fricEnergy, regenElectricalEnergy, battChargeEnergy, regenMechanicalEnergy, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Pneumatic Brake System", recs, evidence);
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

function plotFile = localPlotOverview(outputFolder, t, fricForce, fricPower, brkPedal, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, fricForce, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Friction Brake Force');
ylabel('Force (N)');
grid on;

subplot(3, 1, 2);
plot(t, fricPower, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
yline(config.Thresholds.RegenOpportunityBrakePower_kW, '--', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 1.0);
title('Friction Brake Power');
ylabel('Power (kW)');
legend({'Friction power', 'Regen-opportunity threshold'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
plot(t, brkPedal, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
title('Brake Pedal');
xlabel('Time (s)');
ylabel('Pedal (%)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PneumaticBrakeSystem_Overview', config));
close(fig);
end

function plotFile = localPlotBrakeBlending(outputFolder, t, fricPower, regenElec, battCharge, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, max(fricPower, 0), 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(regenElec, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Friction Braking Versus Electrical Regen');
ylabel('Power (kW)');
legend({'Friction brake power', 'Motor regen electrical power'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, max(regenElec, 0), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, max(battCharge, 0), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Motor Regen Versus Battery Charge Acceptance');
ylabel('Power (kW)');
legend({'Motor regen electrical power', 'Battery charge power'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
brakingMix = NaN(size(fricPower));
valid = (max(fricPower, 0) + max(regenElec, 0)) > 1;
brakingMix(valid) = 100 * max(fricPower(valid), 0) ./ max(max(fricPower(valid), 0) + max(regenElec(valid), 0), eps);
plot(t, brakingMix, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Friction Share of Braking');
xlabel('Time (s)');
ylabel('Friction share (%)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PneumaticBrakeSystem_BrakeBlending', config));
close(fig);
end

function plotFile = localPlotBrakeContext(outputFolder, vehSpeed, roadSlope, fricForce, fricPower, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid1 = isfinite(vehSpeed) & isfinite(fricPower);
scatter(vehSpeed(valid1), fricPower(valid1), 12, config.Plot.Colors.Warning, 'filled');
title('Friction Brake Power Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Friction brake power (kW)');
grid on;

subplot(2, 2, 2);
valid2 = isfinite(roadSlope) & isfinite(fricPower);
scatter(roadSlope(valid2), fricPower(valid2), 12, config.Plot.Colors.Gear, 'filled');
title('Friction Brake Power Versus Road Slope');
xlabel('Road slope (%)');
ylabel('Friction brake power (kW)');
grid on;

subplot(2, 2, 3);
valid3 = isfinite(vehSpeed) & isfinite(fricForce);
scatter(vehSpeed(valid3), abs(fricForce(valid3)), 12, config.Plot.Colors.Vehicle, 'filled');
title('Friction Brake Force Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Brake force magnitude (N)');
grid on;

subplot(2, 2, 4);
histogram(max(fricPower(isfinite(fricPower)), 0), 40, 'FaceColor', config.Plot.Colors.Warning);
title('Friction Brake Power Distribution');
xlabel('Friction brake power (kW)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PneumaticBrakeSystem_BrakeContext', config));
close(fig);
end

function plotFile = localPlotEnergyBreakdown(outputFolder, fricEnergy, regenElectricalEnergy, battChargeEnergy, regenMechanicalEnergy, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

bar(categorical({'Friction loss', 'Motor regen elec', 'Battery charge', 'Motor regen mech'}), ...
    [fricEnergy, regenElectricalEnergy, battChargeEnergy, regenMechanicalEnergy], ...
    'FaceColor', config.Plot.Colors.Vehicle);
title('Braking Energy Breakdown');
ylabel('Energy (kWh)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'PneumaticBrakeSystem_EnergyBreakdown', config));
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
