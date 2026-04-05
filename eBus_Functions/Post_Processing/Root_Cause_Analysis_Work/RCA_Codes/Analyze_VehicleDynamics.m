function result = Analyze_VehicleDynamics(analysisData, outputPaths, config)
% Analyze_VehicleDynamics  Force balance, motion response, and road-load RCA.

result = localInitResult("VEHICLE DYNAMICS", ...
    {'veh_vel', 'veh_acc'}, ...
    {'net_trac_trq', 'fric_brk_force', 'whl_force', 'roll_res_force', 'grad_force', 'aero_drag_force', 'veh_pos'});

d = analysisData.Derived;
t = d.time_s(:);
n = numel(t);
rows = cell(0, 7);
summary = strings(0, 1);
recs = strings(0, 1);
evidence = strings(0, 1);
plotFiles = strings(0, 1);

if isempty(t)
    result.Warnings(end + 1) = "Vehicle dynamics analysis skipped because the aligned time base is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Vehicle Dynamics", recs, evidence);
    result.SummaryText = summary;
    return;
end

vehVel = d.vehVel_kmh(:);
vehAcc = d.vehicleAcceleration_mps2(:);
vehPos = d.vehiclePosition_km(:);
tractionForce = d.tractionForce_N(:);
fricBrakeForce = d.frictionBrakeForce_N(:);
wheelForce = d.wheelForce_N(:);
rollForce = d.rollingResistanceForce_N(:);
gradeForce = d.gradeForce_N(:);
aeroForce = d.aeroDragForce_N(:);
tractionPower = d.tractionPower_kW(:);
roadSlope = d.roadSlope_pct(:);

if ~any(isfinite(vehVel))
    vehVel = localAlignedSignal(analysisData.Signals, 'veh_vel', n);
end
if ~any(isfinite(vehAcc))
    vehAcc = localAlignedSignal(analysisData.Signals, 'veh_acc', n);
end
if ~any(isfinite(vehPos))
    vehPos = localAlignedSignal(analysisData.Signals, 'veh_pos', n);
end
if ~any(isfinite(tractionForce))
    tractionForce = localAlignedSignal(analysisData.Signals, 'net_trac_trq', n);
end
if ~any(isfinite(fricBrakeForce))
    fricBrakeForce = localAlignedSignal(analysisData.Signals, 'fric_brk_force', n);
end
if ~any(isfinite(wheelForce))
    wheelForce = localAlignedSignal(analysisData.Signals, 'whl_force', n);
end
if ~any(isfinite(rollForce))
    rollForce = localAlignedSignal(analysisData.Signals, 'roll_res_force', n);
end
if ~any(isfinite(gradeForce))
    gradeForce = localAlignedSignal(analysisData.Signals, 'grad_force', n);
end
if ~any(isfinite(aeroForce))
    aeroForce = localAlignedSignal(analysisData.Signals, 'aero_drag_force', n);
end

if all(isnan(wheelForce)) && ~all(isnan(tractionForce))
    wheelForce = tractionForce;
end

if all(isnan(vehVel))
    result.Warnings(end + 1) = "Vehicle dynamics analysis skipped because vehicle speed is unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Vehicle Dynamics", recs, evidence);
    result.SummaryText = summary;
    return;
end

vehVelMps = vehVel / 3.6;
tripDistance = d.tripDistance_km;
movingMask = isfinite(vehVel) & vehVel > config.Thresholds.StopSpeed_kmh;
accelMask = isfinite(vehAcc) & vehAcc >= config.Thresholds.SignificantAccel_mps2;
decelMask = isfinite(vehAcc) & vehAcc <= config.Thresholds.SignificantDecel_mps2;
cruiseMask = movingMask & isfinite(vehAcc) & abs(vehAcc) <= config.Thresholds.CruiseAccelAbs_mps2;
uphillMask = movingMask & isfinite(roadSlope) & roadSlope >= config.Thresholds.UphillSlope_pct;
steepUphillMask = movingMask & isfinite(roadSlope) & roadSlope >= config.Thresholds.SteepSlope_pct;

roadLoad = rollForce + gradeForce + aeroForce;
netLongitudinalForce = tractionForce - abs(fricBrakeForce) - roadLoad;
forceBalanceError = wheelForce - (tractionForce - abs(fricBrakeForce));

roadLoadEnergy = RCA_TrapzFinite(t, max(d.resistivePower_kW(:), 0)) / 3600;
tractiveEnergy = RCA_TrapzFinite(t, max(tractionPower, 0)) / 3600;

rows = RCA_AddKPI(rows, 'Trip Distance', tripDistance, 'km', ...
    'Operation', 'Vehicle Dynamics', 'veh_pos or integrated veh_vel', ...
    'Vehicle distance basis selected from logged position or integrated speed.');
rows = RCA_AddKPI(rows, 'Average Vehicle Speed', mean(vehVel, 'omitnan'), 'km/h', ...
    'Operation', 'Vehicle Dynamics', 'veh_vel', ...
    'Average vehicle speed over the full trip.');
rows = RCA_AddKPI(rows, 'Peak Vehicle Speed', max(vehVel, [], 'omitnan'), 'km/h', ...
    'Performance', 'Vehicle Dynamics', 'veh_vel', ...
    'Maximum observed vehicle speed.');
rows = RCA_AddKPI(rows, 'Peak Vehicle Acceleration', max(vehAcc, [], 'omitnan'), 'm/s^2', ...
    'Performance', 'Vehicle Dynamics', 'veh_acc', ...
    'Maximum positive vehicle acceleration.');
rows = RCA_AddKPI(rows, 'Peak Vehicle Deceleration', min(vehAcc, [], 'omitnan'), 'm/s^2', ...
    'Performance', 'Vehicle Dynamics', 'veh_acc', ...
    'Maximum negative vehicle acceleration.');
rows = RCA_AddKPI(rows, 'Acceleration Event Share', 100 * RCA_FractionTrue(accelMask, movingMask), '%', ...
    'Operation', 'Vehicle Dynamics', 'veh_acc + veh_vel', ...
    'Share of moving samples classified as acceleration events.');
rows = RCA_AddKPI(rows, 'Deceleration Event Share', 100 * RCA_FractionTrue(decelMask, movingMask), '%', ...
    'Operation', 'Vehicle Dynamics', 'veh_acc + veh_vel', ...
    'Share of moving samples classified as deceleration events.');
rows = RCA_AddKPI(rows, 'Cruise Event Share', 100 * RCA_FractionTrue(cruiseMask, movingMask), '%', ...
    'Operation', 'Vehicle Dynamics', 'veh_acc + veh_vel', ...
    'Share of moving samples classified as cruise or steady-speed operation.');
rows = RCA_AddKPI(rows, 'Integrated Resistive Road-Load Energy', roadLoadEnergy, 'kWh', ...
    'Losses', 'Vehicle Dynamics', 'rolling + grade + aero force with speed', ...
    'Integrated positive road-load power.');
rows = RCA_AddKPI(rows, 'Mean Rolling Resistance Force', mean(rollForce, 'omitnan'), 'N', ...
    'Losses', 'Vehicle Dynamics', 'roll_res_force', ...
    'Average rolling resistance force.');
rows = RCA_AddKPI(rows, 'Mean Gradient Force', mean(gradeForce, 'omitnan'), 'N', ...
    'Operation', 'Vehicle Dynamics', 'grad_force', ...
    'Average grade force over the trip.');
rows = RCA_AddKPI(rows, 'Mean Aerodynamic Force', mean(aeroForce, 'omitnan'), 'N', ...
    'Losses', 'Vehicle Dynamics', 'aero_drag_force', ...
    'Average aerodynamic drag force.');
rows = RCA_AddKPI(rows, 'Road-Load Share: Rolling', 100 * mean(max(rollForce, 0), 'omitnan') / max(mean(max(roadLoad, 0), 'omitnan'), eps), '%', ...
    'Losses', 'Vehicle Dynamics', 'rolling resistance / total road load', ...
    'Share of average positive road load due to rolling resistance.');
rows = RCA_AddKPI(rows, 'Road-Load Share: Grade', 100 * mean(max(gradeForce, 0), 'omitnan') / max(mean(max(roadLoad, 0), 'omitnan'), eps), '%', ...
    'Losses', 'Vehicle Dynamics', 'grade force / total road load', ...
    'Share of average positive road load due to uphill grade.');
rows = RCA_AddKPI(rows, 'Road-Load Share: Aero', 100 * mean(max(aeroForce, 0), 'omitnan') / max(mean(max(roadLoad, 0), 'omitnan'), eps), '%', ...
    'Losses', 'Vehicle Dynamics', 'aero drag / total road load', ...
    'Share of average positive road load due to aerodynamic drag.');
rows = RCA_AddKPI(rows, 'Mean Net Longitudinal Force', mean(netLongitudinalForce(movingMask), 'omitnan'), 'N', ...
    'Balance', 'Vehicle Dynamics', 'traction - friction - road load', ...
    'Average remaining longitudinal force after braking and road-load subtraction.');
rows = RCA_AddKPI(rows, 'Force Balance Error', mean(abs(forceBalanceError), 'omitnan'), 'N', ...
    'Consistency', 'Vehicle Dynamics', 'wheel force versus tractive and brake force balance', ...
    'Mean mismatch between wheel force and tractive-minus-brake force.');
rows = RCA_AddKPI(rows, 'Uphill Driving Share', 100 * RCA_FractionTrue(uphillMask, movingMask), '%', ...
    'Operation', 'Vehicle Dynamics', 'veh_vel + road_slp', ...
    'Share of moving samples in uphill operation.');
rows = RCA_AddKPI(rows, 'Steep Uphill Share', 100 * RCA_FractionTrue(steepUphillMask, movingMask), '%', ...
    'Operation', 'Vehicle Dynamics', 'veh_vel + road_slp', ...
    'Share of moving samples in steep uphill operation.');
rows = RCA_AddKPI(rows, 'Mean Uphill Net Longitudinal Force', mean(netLongitudinalForce(uphillMask), 'omitnan'), 'N', ...
    'Capability', 'Vehicle Dynamics', 'force balance + road slope', ...
    'Average remaining longitudinal force in uphill operation.');
rows = RCA_AddKPI(rows, 'Positive Net Force Share', 100 * RCA_FractionTrue(netLongitudinalForce > 0, movingMask), '%', ...
    'Capability', 'Vehicle Dynamics', 'net longitudinal force', ...
    'Share of moving samples with positive residual longitudinal force.');

summary(end + 1) = sprintf(['Vehicle dynamics summary: %.2f km trip distance, %.1f km/h average vehicle speed, and %.2f kWh integrated resistive road-load energy.'], ...
    tripDistance, mean(vehVel, 'omitnan'), roadLoadEnergy);
summary(end + 1) = sprintf(['Road-load split: rolling %.0f N, grade %.0f N, and aero %.0f N on average. ', ...
    'This identifies the dominant external force contributors over the route.'], ...
    mean(rollForce, 'omitnan'), mean(gradeForce, 'omitnan'), mean(aeroForce, 'omitnan'));
summary(end + 1) = sprintf(['Force-balance context: mean force-balance error is %.0f N and mean uphill net longitudinal force is %.0f N. ', ...
    'This indicates whether delivered wheel force is consistent with modeled vehicle motion demand.'], ...
    mean(abs(forceBalanceError), 'omitnan'), mean(netLongitudinalForce(uphillMask), 'omitnan'));

if mean(rollForce, 'omitnan') > mean(aeroForce, 'omitnan') * 1.5
    recs(end + 1) = "Rolling resistance dominates aerodynamic drag; review tyre, road-loss, or wheel-loss assumptions before attributing poor efficiency to aero effects.";
    evidence(end + 1) = "Mean rolling resistance force materially exceeds mean aerodynamic force.";
end
if mean(max(gradeForce, 0), 'omitnan') > 0.25 * mean(max(tractionForce, 0), 'omitnan')
    recs(end + 1) = "Route grade is a first-order vehicle-dynamics driver; keep grade-normalized comparisons separate from flat-route performance discussions.";
    evidence(end + 1) = "Mean positive grade force is a large fraction of delivered tractive force.";
end
if mean(abs(forceBalanceError), 'omitnan') > 500
    recs(end + 1) = "Check wheel-force, tractive-force, and brake-force consistency because the longitudinal force balance shows material mismatch.";
    evidence(end + 1) = sprintf('Mean force-balance error is %.0f N.', mean(abs(forceBalanceError), 'omitnan'));
end
if mean(netLongitudinalForce(uphillMask), 'omitnan') < 0
    recs(end + 1) = "Investigate uphill performance limitation. Net longitudinal force is negative on average in uphill operation, so delivered force may be insufficient to sustain demanded motion.";
    evidence(end + 1) = sprintf('Mean uphill net longitudinal force is %.0f N.', mean(netLongitudinalForce(uphillMask), 'omitnan'));
end

figureFolder = fullfile(outputPaths.FiguresSubsystem, 'VehicleDynamics');
plotFiles = localAppendPlotFile(plotFiles, localPlotForceBreakdown(figureFolder, t, rollForce, gradeForce, aeroForce, tractionForce, fricBrakeForce, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotMotionResponse(figureFolder, t, vehVel, vehAcc, vehPos, netLongitudinalForce, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotRoadLoadContext(figureFolder, vehVel, roadSlope, roadLoad, tractionForce, config));
plotFiles = localAppendPlotFile(plotFiles, localPlotForceBalance(figureFolder, t, wheelForce, tractionForce, fricBrakeForce, forceBalanceError, config));
plotFiles = plotFiles(plotFiles ~= "");

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Vehicle Dynamics", recs, evidence);
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

function plotFile = localPlotForceBreakdown(outputFolder, t, rollForce, gradeForce, aeroForce, tractionForce, fricBrakeForce, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, rollForce, 'Color', config.Plot.Colors.Neutral, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, gradeForce, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
plot(t, aeroForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Road-Load Force Breakdown');
ylabel('Force (N)');
legend({'Rolling', 'Grade', 'Aero'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, tractionForce, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, abs(fricBrakeForce), 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
title('Traction and Friction Brake Force');
ylabel('Force (N)');
legend({'Net tractive force', 'Friction brake force magnitude'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
plot(t, tractionForce - abs(fricBrakeForce), 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
title('Delivered Longitudinal Force Before Road Load');
xlabel('Time (s)');
ylabel('Force (N)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'VehicleDynamics_ForceBreakdown', config));
close(fig);
end

function plotFile = localPlotMotionResponse(outputFolder, t, vehVel, vehAcc, vehPos, netLongitudinalForce, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(4, 1, 1);
plot(t, vehVel, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Vehicle Velocity');
ylabel('Speed (km/h)');
grid on;

subplot(4, 1, 2);
plot(t, vehAcc, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Vehicle Acceleration');
ylabel('Acc (m/s^2)');
grid on;

subplot(4, 1, 3);
plot(t, vehPos, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
title('Vehicle Position');
ylabel('Position (km)');
grid on;

subplot(4, 1, 4);
plot(t, netLongitudinalForce, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Net Longitudinal Force');
xlabel('Time (s)');
ylabel('Force (N)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'VehicleDynamics_MotionResponse', config));
close(fig);
end

function plotFile = localPlotRoadLoadContext(outputFolder, vehVel, roadSlope, roadLoad, tractionForce, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(2, 2, 1);
valid1 = isfinite(vehVel) & isfinite(roadLoad);
scatter(vehVel(valid1), roadLoad(valid1), 12, config.Plot.Colors.Warning, 'filled');
title('Road Load Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Road load (N)');
grid on;

subplot(2, 2, 2);
valid2 = isfinite(roadSlope) & isfinite(roadLoad);
scatter(roadSlope(valid2), roadLoad(valid2), 12, config.Plot.Colors.Slope, 'filled');
title('Road Load Versus Road Slope');
xlabel('Road slope (%)');
ylabel('Road load (N)');
grid on;

subplot(2, 2, 3);
valid3 = isfinite(vehVel) & isfinite(tractionForce);
scatter(vehVel(valid3), tractionForce(valid3), 12, roadSlope(valid3), 'filled');
title('Tractive Force Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('Tractive force (N)');
cb = colorbar;
cb.Label.String = 'Road slope (%)';
grid on;

subplot(2, 2, 4);
valid4 = isfinite(roadLoad) & isfinite(tractionForce);
scatter(roadLoad(valid4), tractionForce(valid4), 12, config.Plot.Colors.Vehicle, 'filled');
title('Tractive Force Versus Road Load');
xlabel('Road load (N)');
ylabel('Tractive force (N)');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'VehicleDynamics_RoadLoadContext', config));
close(fig);
end

function plotFile = localPlotForceBalance(outputFolder, t, wheelForce, tractionForce, fricBrakeForce, forceBalanceError, config)
plotFile = "";
fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);

subplot(3, 1, 1);
plot(t, wheelForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, tractionForce - abs(fricBrakeForce), 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
title('Wheel Force Versus Tractive-Minus-Brake Force');
ylabel('Force (N)');
legend({'Wheel force', 'Tractive - friction brake'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, forceBalanceError, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Force-Balance Error');
ylabel('Error (N)');
grid on;

subplot(3, 1, 3);
histogram(forceBalanceError(isfinite(forceBalanceError)), 40, 'FaceColor', config.Plot.Colors.Warning);
title('Force-Balance Error Distribution');
xlabel('Error (N)');
ylabel('Samples');
grid on;

plotFile = string(RCA_SaveFigure(fig, outputFolder, 'VehicleDynamics_ForceBalance', config));
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
