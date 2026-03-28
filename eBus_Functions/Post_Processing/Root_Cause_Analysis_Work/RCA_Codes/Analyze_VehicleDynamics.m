function result = Analyze_VehicleDynamics(analysisData, outputPaths, config)
% Analyze_VehicleDynamics  Road-load and motion response analysis.

result = localInitResult("VEHICLE DYNAMICS", {'veh_vel', 'veh_acc'}, {'roll_res_force', 'grad_force', 'aero_drag_force', 'whl_force'});
t = analysisData.Derived.time_s;
rows = cell(0, 7);
summary = strings(0, 1);

rows = RCA_AddKPI(rows, 'Trip Distance', analysisData.Derived.tripDistance_km, 'km', 'Operation', 'Vehicle Dynamics', 'veh_pos or integrated veh_vel', 'Vehicle distance basis is automatically selected.');
rows = RCA_AddKPI(rows, 'Average Vehicle Speed', mean(analysisData.Derived.vehVel_kmh, 'omitnan'), 'km/h', 'Operation', 'Vehicle Dynamics', 'veh_vel', 'Complete if vehicle speed is available.');
rows = RCA_AddKPI(rows, 'Peak Vehicle Acceleration', max(analysisData.Derived.vehicleAcceleration_mps2, [], 'omitnan'), 'm/s^2', 'Performance', 'Vehicle Dynamics', 'veh_acc or differentiated veh_vel', 'Acceleration is derived if direct logging is missing.');

resistiveEnergy = trapz(t, max(analysisData.Derived.resistivePower_kW, 0)) / 3600;
rows = RCA_AddKPI(rows, 'Integrated Resistive Road-Load Energy', resistiveEnergy, 'kWh', 'Losses', 'Vehicle Dynamics', 'rolling + grade + aero force with speed', 'Physical road-load basis.');
rows = RCA_AddKPI(rows, 'Mean Rolling Resistance Force', mean(analysisData.Derived.rollingResistanceForce_N, 'omitnan'), 'N', 'Losses', 'Vehicle Dynamics', 'roll_res_force', 'Available if rolling resistance is logged.');
rows = RCA_AddKPI(rows, 'Mean Gradient Force', mean(analysisData.Derived.gradeForce_N, 'omitnan'), 'N', 'Operation', 'Vehicle Dynamics', 'grad_force', 'Available if grade force is logged.');
rows = RCA_AddKPI(rows, 'Mean Aerodynamic Force', mean(analysisData.Derived.aeroDragForce_N, 'omitnan'), 'N', 'Losses', 'Vehicle Dynamics', 'aero_drag_force', 'Available if aero drag is logged.');
summary(end + 1) = sprintf('Vehicle dynamics summary: integrated resistive road-load energy is %.2f kWh across %.2f km.', ...
    resistiveEnergy, analysisData.Derived.tripDistance_km);

recs = strings(0, 1);
evidence = strings(0, 1);
if mean(analysisData.Derived.rollingResistanceForce_N, 'omitnan') > mean(analysisData.Derived.aeroDragForce_N, 'omitnan') * 1.5
    recs(end + 1) = "Rolling resistance dominates aero drag; tyre, wheel-loss, or road-loss assumptions deserve a focused review.";
    evidence(end + 1) = "Mean rolling resistance force materially exceeds mean aerodynamic force.";
end
if mean(analysisData.Derived.gradeForce_N, 'omitnan') > 0.25 * mean(max(analysisData.Derived.tractionForce_N, 0), 'omitnan')
    recs(end + 1) = "Route grade is a first-order driver; keep grade-normalized comparisons separate from flat-route efficiency discussions.";
    evidence(end + 1) = "Grade force is a large fraction of delivered traction force.";
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, analysisData.Derived.rollingResistanceForce_N, 'Color', config.Plot.Colors.Neutral, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, analysisData.Derived.gradeForce_N, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
plot(t, analysisData.Derived.aeroDragForce_N, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
legend({'Rolling', 'Grade', 'Aero'}, 'Location', 'best');
title('Vehicle Dynamics Force Breakdown');
ylabel('Force (N)');
grid on;

subplot(2, 1, 2);
plot(t, analysisData.Derived.vehVel_kmh, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, analysisData.Derived.vehicleAcceleration_mps2 * 10, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
title('Vehicle Speed and Acceleration Trend');
xlabel('Time (s)');
ylabel('Speed (km/h), Acc x10');
legend({'Vehicle speed', 'Acceleration x10'}, 'Location', 'best');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'VehicleDynamics'), 'VehicleDynamics_Forces', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Vehicle Dynamics", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
