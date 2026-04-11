function plotResults = RCA_GenerateVehiclePlots(analysisData, outputPaths, config)
% RCA_GenerateVehiclePlots  Create vehicle-level engineering plots.

derived = analysisData.Derived;
t = derived.time_s;
plotFiles = strings(0, 1);
plotNotes = strings(0, 1);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
yyaxis left;
hSpeed = plot(t, derived.vehVel_kmh, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
legendHandles = hSpeed;
legendLabels = {'Vehicle speed'};
if ~all(isnan(derived.speedDemand_kmh))
    hDemand = plot(t, derived.speedDemand_kmh, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    legendHandles(end + 1) = hDemand; %#ok<AGROW>
    legendLabels{end + 1} = 'Desired speed'; %#ok<AGROW>
end
ylabel('Speed (km/h)');
if isfield(derived, 'roadSlope_pct') && ~all(isnan(derived.roadSlope_pct))
    yyaxis right;
    hSlope = plot(t, derived.roadSlope_pct, ':', 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
    ylabel('Road slope (%)');
    legendHandles(end + 1) = hSlope; %#ok<AGROW>
    legendLabels{end + 1} = 'Road slope'; %#ok<AGROW>
    yyaxis left;
end
title('Vehicle Speed Tracking');
legend(legendHandles, legendLabels, 'Location', 'best');
grid on;

subplot(2, 1, 2);
yyaxis left;
hAcc = plot(t, derived.vehicleAcceleration_mps2, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
ylabel('Acceleration (m/s^2)');
legendHandles = hAcc;
legendLabels = {'Acceleration'};
if isfield(derived, 'gearNumber') && ~all(isnan(derived.gearNumber))
    yyaxis right;
    hGear = stairs(t, derived.gearNumber, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
    ylabel('Current gear (-)');
    legendHandles(end + 1) = hGear; %#ok<AGROW>
    legendLabels{end + 1} = 'Current gear'; %#ok<AGROW>
    yyaxis left;
end
title('Acceleration and Current Gear');
xlabel('Time (s)');
legend(legendHandles, legendLabels, 'Location', 'best');
grid on;
plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Speed_Tracking', config));
plotNotes(end + 1) = "Speed tracking plot compares demanded and delivered vehicle response and contextualizes response with road slope, acceleration, and current gear.";
close(fig);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(3, 1, 1);
plot(t, derived.batteryPower_kW, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.auxiliaryPower_kW, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
if ~all(isnan(derived.highPowerResistorPower_kW))
    plot(t, derived.highPowerResistorPower_kW, '-.', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
end
plot(t, derived.motorLossPower_kW + derived.gearboxLossPower_kW + derived.batteryLossPower_kW, ':', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Vehicle Power Overview (Battery Discharge +, Charge -)');
ylabel('Power (kW)');
legendEntries = {'Battery power (RCA sign)', 'Auxiliary power'};
if ~all(isnan(derived.highPowerResistorPower_kW))
    legendEntries{end + 1} = 'High-power resistor';
end
legendEntries{end + 1} = 'Logged loss power';
legendEntries{end + 1} = 'Zero line';
legend(legendEntries, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, RCA_CumtrapzFinite(t, max(derived.batteryPower_kW, 0)) / 3600, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, RCA_CumtrapzFinite(t, max(-derived.batteryPower_kW, 0)) / 3600, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, RCA_CumtrapzFinite(t, derived.batteryPower_kW) / 3600, '-.', 'Color', config.Plot.Colors.Neutral, 'LineWidth', config.Plot.LineWidth);
plot(t, RCA_CumtrapzFinite(t, max(derived.auxiliaryPower_kW, 0)) / 3600, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
plot(t, RCA_CumtrapzFinite(t, max(derived.tractionPower_kW, 0)) / 3600, ':', 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Cumulative Energy');
ylabel('Energy (kWh)');
legend({'Battery discharge', 'Battery recovery', 'Net battery energy', 'Auxiliary', 'Traction'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
bar(categorical({'Battery loss', 'Motor loss', 'Transmission loss', 'Friction brake', 'Auxiliary'}), ...
    [RCA_TrapzFinite(t, max(derived.batteryLossPower_kW, 0)) / 3600, ...
    RCA_TrapzFinite(t, max(derived.motorLossPower_kW, 0)) / 3600, ...
    RCA_TrapzFinite(t, max(derived.gearboxLossPower_kW, 0)) / 3600, ...
    RCA_TrapzFinite(t, max(derived.frictionBrakePower_kW, 0)) / 3600, ...
    RCA_TrapzFinite(t, max(derived.auxiliaryPower_kW, 0)) / 3600], 'FaceColor', config.Plot.Colors.Vehicle);
title('Integrated Loss and Auxiliary Breakdown');
ylabel('Energy (kWh)');
grid on;
plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Energy_Overview', config));
plotNotes(end + 1) = "Energy overview plot summarizes the electrical burden, cumulative energy usage, and logged loss contributors. Battery power is normalized to discharge-positive sign from workbook metadata.";
close(fig);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(3, 1, 1);
plot(t, derived.batteryPower_kW, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.dcBusLoadPower_kW, '--', 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
plot(t, derived.batteryPower_kW - derived.batteryLossPower_kW, ':', 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
title('Vehicle Power Balance');
ylabel('Power (kW)');
legend({'Battery power', 'Motor + auxiliary + HPR', 'Battery power minus battery loss'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, derived.powerBalanceResidualTerminal_kW, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.powerBalanceResidualInternal_kW, '--', 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, config.Thresholds.PowerBalanceResidualWarn_kW * ones(size(t)), ':', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
plot(t, -config.Thresholds.PowerBalanceResidualWarn_kW * ones(size(t)), ':', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Power Balance Residuals');
ylabel('Residual (kW)');
legend({'Terminal residual', 'Residual with battery loss', 'Warning threshold'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
scatter(derived.vehVel_kmh, abs(derived.powerBalanceResidualTerminal_kW), 12, derived.roadSlope_pct, 'filled');
title('Power Balance Residual Versus Vehicle Speed');
xlabel('Vehicle speed (km/h)');
ylabel('|Terminal residual| (kW)');
cb = colorbar;
cb.Label.String = 'Road slope (%)';
grid on;
plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Power_Balance', config));
plotNotes(end + 1) = "Power-balance plot compares battery source power against motor, auxiliary, and resistor sink power and shows the residual mismatch over the trip.";
close(fig);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(3, 1, 1);
plot(t, derived.vehiclePropulsionForce_N, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.frictionBrakeForce_N, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, derived.wheelForce_N, ':', 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Vehicle Force Path');
ylabel('Force (N)');
legend({'Vehicle propulsion force', 'Friction brake force', 'Wheel force'}, 'Location', 'best');
grid on;

subplot(3, 1, 2);
plot(t, derived.roadLoadForce_N, 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.inertialForce_N, '--', 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
plot(t, derived.wheelForce_N, ':', 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
title('Wheel Force Versus Road-Load Balance');
ylabel('Force (N)');
legend({'Rolling + grade + aero', 'Inertial force', 'Wheel force'}, 'Location', 'best');
grid on;

subplot(3, 1, 3);
plot(t, derived.forceBalanceResidualWheel_N, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.forceBalanceResidualRoadLoad_N, '--', 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
plot(t, config.Thresholds.ForceBalanceResidualWarn_N * ones(size(t)), ':', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
plot(t, -config.Thresholds.ForceBalanceResidualWarn_N * ones(size(t)), ':', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Force Balance Residuals');
xlabel('Time (s)');
ylabel('Residual (N)');
legend({'Wheel-path residual', 'Road-load residual', 'Warning threshold'}, 'Location', 'best');
grid on;
plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Force_Balance', config));
plotNotes(end + 1) = "Force-balance plot compares propulsion, braking, wheel, road-load, and inertial forces and exposes residual mismatch through the trip.";
close(fig);

energyData = localComputeVehicleEnergyFlow(derived, t);
energyFigurePosition = config.Plot.FigurePosition;
energyFigurePosition(3:4) = [1500 900];
fig = figure('Color', 'w', 'Position', energyFigurePosition);
ax = axes('Parent', fig, 'Position', [0.02 0.04 0.96 0.92]);
axis(ax, [0 1 0 1]);
axis(ax, 'off');
hold(ax, 'on');

localDrawEnergyNode(ax, [0.05 0.50 0.19 0.13], sprintf('Battery\nDischarge\n%.2f kWh', energyData.Discharge_kWh), config.Plot.Colors.Battery, 12);
localDrawEnergyNode(ax, [0.31 0.50 0.21 0.13], sprintf('DC Bus\nAvailable\n%.2f kWh', energyData.NetBus_kWh), config.Plot.Colors.Vehicle, 12);
localDrawEnergyNode(ax, [0.59 0.50 0.18 0.13], sprintf('Wheel\nTraction\n%.2f kWh', energyData.Traction_kWh), config.Plot.Colors.Motor, 12);
localDrawEnergyNode(ax, [0.82 0.50 0.13 0.13], sprintf('Distance\n%.2f km', derived.tripDistance_km), config.Plot.Colors.Demand, 12);

localDrawEnergyNode(ax, [0.09 0.73 0.14 0.10], sprintf('Battery Loss\n%.2f kWh', energyData.BattLoss_kWh), config.Plot.Colors.Warning, 10);
localDrawEnergyNode(ax, [0.31 0.73 0.15 0.10], sprintf('Auxiliaries\n%.2f kWh', energyData.Aux_kWh), config.Plot.Colors.Auxiliary, 10);
if energyData.HPR_kWh > 0
    localDrawEnergyNode(ax, [0.46 0.73 0.10 0.10], sprintf('HPR\n%.2f kWh', energyData.HPR_kWh), config.Plot.Colors.Demand, 10);
end
localDrawEnergyNode(ax, [0.58 0.73 0.16 0.10], sprintf('Motor / Inverter Loss\n%.2f kWh', energyData.MotorLoss_kWh), config.Plot.Colors.Warning, 10);
localDrawEnergyNode(ax, [0.79 0.73 0.16 0.10], sprintf('Transmission Loss\n%.2f kWh', energyData.GbxLoss_kWh), config.Plot.Colors.Warning, 10);

localDrawEnergyNode(ax, [0.39 0.25 0.22 0.11], sprintf('Braking Energy Split\n%.2f kWh', energyData.BrakeSplit_kWh), config.Plot.Colors.Neutral, 11);
localDrawEnergyNode(ax, [0.16 0.05 0.20 0.11], sprintf('Battery Regen\n%.2f kWh', energyData.Regen_kWh), config.Plot.Colors.Battery, 11);
localDrawEnergyNode(ax, [0.65 0.05 0.20 0.11], sprintf('Friction Brake\n%.2f kWh', energyData.Friction_kWh), config.Plot.Colors.Warning, 11);

localDrawEnergyArrow(ax, [0.24 0.565], [0.31 0.565], sprintf('%.2f kWh', energyData.Discharge_kWh), [0.00 0.03]);
localDrawEnergyArrow(ax, [0.52 0.565], [0.59 0.565], sprintf('%.2f kWh', energyData.Traction_kWh), [0.00 0.03]);
localDrawEnergyArrow(ax, [0.77 0.565], [0.82 0.565], 'Vehicle motion', [0.00 0.03]);
localDrawEnergyArrow(ax, [0.16 0.63], [0.16 0.73], sprintf('%.2f', energyData.BattLoss_kWh), [-0.03 0.00]);
localDrawEnergyArrow(ax, [0.38 0.63], [0.38 0.73], sprintf('%.2f', energyData.Aux_kWh), [0.03 0.00]);
if energyData.HPR_kWh > 0
    localDrawEnergyArrow(ax, [0.51 0.63], [0.51 0.73], sprintf('%.2f', energyData.HPR_kWh), [0.03 0.00]);
end
localDrawEnergyArrow(ax, [0.66 0.63], [0.66 0.73], sprintf('%.2f', energyData.MotorLoss_kWh), [0.03 0.00]);
localDrawEnergyArrow(ax, [0.87 0.63], [0.87 0.73], sprintf('%.2f', energyData.GbxLoss_kWh), [0.03 0.00]);
localDrawEnergyArrow(ax, [0.66 0.50], [0.50 0.36], sprintf('Braking domain\n%.2f kWh', energyData.BrakeSplit_kWh), [0.00 0.04]);
localDrawEnergyArrow(ax, [0.46 0.25], [0.30 0.16], sprintf('Recovered\n%.2f kWh', energyData.Regen_kWh), [0.00 0.03]);
localDrawEnergyArrow(ax, [0.54 0.25], [0.74 0.16], sprintf('Dissipated\n%.2f kWh', energyData.Friction_kWh), [0.00 0.03]);

text(ax, 0.50, 0.95, 'Vehicle Energy Flow Diagram', 'HorizontalAlignment', 'center', ...
    'FontWeight', 'bold', 'FontSize', 15);
text(ax, 0.50, 0.91, sprintf('Trip net battery energy %.2f kWh | Battery-to-wheel efficiency %.1f%% | Regen recovery %.1f%%', ...
    energyData.NetBattery_kWh, energyData.BatteryToWheelEff_pct, energyData.RegenRecovery_pct), ...
    'HorizontalAlignment', 'center', 'FontSize', 11, 'FontWeight', 'bold');
text(ax, 0.50, 0.87, 'RCA sign convention: battery discharge positive, battery charge / regeneration negative in workbook source data.', ...
    'HorizontalAlignment', 'center', 'FontSize', 9, 'Color', [0.20 0.20 0.20]);

plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Energy_Flow_Diagram', config));
plotNotes(end + 1) = "Energy flow diagram summarizes how battery discharge energy is distributed across auxiliaries, internal losses, wheel traction, and braking recovery or dissipation.";
close(fig);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 2, 1);
stairs(t, derived.gearNumber, 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth);
title('Gear Number vs Time');
xlabel('Time (s)');
ylabel('Gear');
grid on;

subplot(2, 2, 2);
scatter(derived.vehVel_kmh, derived.gearNumber, 12, abs(derived.motorSpeed_rpm), 'filled');
title('Gear vs Vehicle Speed');
xlabel('Vehicle Speed (km/h)');
ylabel('Gear');
cb = colorbar;
cb.Label.String = 'Motor speed magnitude (rpm)';
grid on;

subplot(2, 2, 3);
histogram(derived.gearNumber(~isnan(derived.gearNumber)), 'FaceColor', config.Plot.Colors.Gear);
title('Gear Usage Histogram');
xlabel('Gear');
ylabel('Samples');
grid on;

subplot(2, 2, 4);
scatter(abs(derived.motorSpeed_rpm), derived.torqueActualTotal_Nm, 12, derived.gearNumber, 'filled');
title('Motor Operating Region Colored by Gear');
xlabel('Motor speed (rpm)');
ylabel('Motor torque (Nm)');
cb = colorbar;
cb.Label.String = 'Gear';
grid on;
plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Gear_Analysis', config));
plotNotes(end + 1) = "Gear analysis plot shows whether shift logic keeps the motors in stable and efficient operating regions.";
close(fig);

if ~isempty(analysisData.SegmentSummary)
    seg = analysisData.SegmentSummary;
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(2, 1, 1);
    bar(seg.SegmentID, seg.Wh_per_km, 'FaceColor', config.Plot.Colors.Vehicle); hold on;
    poorIdx = seg.IsPoorEfficiency;
    plot(seg.SegmentID(poorIdx), seg.Wh_per_km(poorIdx), 'o', 'Color', config.Plot.Colors.Warning, 'MarkerSize', 8, 'LineWidth', 1.5);
    title('Segment Energy Intensity Ranking');
    xlabel('Segment ID');
    ylabel('Wh/km');
    grid on;

    subplot(2, 1, 2);
    bar(seg.SegmentID, [seg.AuxEnergyShare_pct, seg.LossShare_pct], 'stacked');
    title('Segment Auxiliary and Loss Share');
    xlabel('Segment ID');
    ylabel({'Share of battery discharge (%)', '(net consuming segments only)'});
    legend({'Auxiliary share', 'Loss share'}, 'Location', 'best');
    grid on;
    plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Segment_Ranking', config));
    plotNotes(end + 1) = "Segment ranking plot highlights inefficient segments and separates auxiliary burden from conversion losses.";
    close(fig);
end

if ~isempty(analysisData.RootCauseRanking)
    [causeNames, contributionSums] = localAggregateCauseContributions(analysisData.RootCauseRanking);
    cumulative = cumsum(contributionSums) / max(sum(contributionSums), eps) * 100;

    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    causeIndex = 1:numel(causeNames);
    yyaxis left;
    bar(causeIndex, contributionSums, 'FaceColor', config.Plot.Colors.Warning);
    ylabel('Aggregated contribution (%)');
    yyaxis right;
    plot(causeIndex, cumulative, '-o', 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
    ylabel('Cumulative contribution (%)');
    title('Pareto-Style Root Cause Ranking');
    xlabel('Cause');
    xticks(causeIndex);
    xticklabels(cellstr(causeNames));
    xtickangle(30);
    grid on;
    plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_RootCause_Pareto', config));
    plotNotes(end + 1) = "Pareto-style plot ranks the recurring physical drivers across the worst trip segments.";
    close(fig);
end

worstSegments = analysisData.BadSegmentTable;
numDashboards = min(height(worstSegments), 3);
for iDash = 1:numDashboards
    segID = worstSegments.SegmentID(iDash);
    segRow = analysisData.SegmentSummary(analysisData.SegmentSummary.SegmentID == segID, :);
    idx = segRow.StartIndex:segRow.EndIndex;

    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(4, 1, 1);
    plot(t(idx), derived.vehVel_kmh(idx), 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    if ~all(isnan(derived.speedDemand_kmh(idx)))
        plot(t(idx), derived.speedDemand_kmh(idx), '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    end
    title(sprintf('Worst Segment %d Dashboard', segID));
    ylabel('Speed (km/h)');
    grid on;

    subplot(4, 1, 2);
    plot(t(idx), derived.batteryPower_kW(idx), 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
    plot(t(idx), derived.auxiliaryPower_kW(idx), '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
    ylabel('Power (kW)');
    legend({'Battery (RCA sign)', 'Auxiliary'}, 'Location', 'best');
    grid on;

    subplot(4, 1, 3);
    stairs(t(idx), derived.gearNumber(idx), 'Color', config.Plot.Colors.Gear, 'LineWidth', config.Plot.LineWidth); hold on;
    plot(t(idx), abs(derived.motorSpeed_rpm(idx)) / max(max(abs(derived.motorSpeed_rpm), [], 'omitnan'), 1) * max(max(derived.gearNumber, [], 'omitnan'), 1), '--', ...
        'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
    ylabel('Gear / norm speed');
    legend({'Gear', 'Normalized motor speed'}, 'Location', 'best');
    grid on;

    subplot(4, 1, 4);
    plot(t(idx), derived.roadSlope_pct(idx), 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth); hold on;
    plot(t(idx), derived.frictionBrakePower_kW(idx), '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
    xlabel('Time (s)');
    ylabel('Slope / brake');
    legend({'Slope (%)', 'Friction brake power (kW)'}, 'Location', 'best');
    grid on;

    plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, sprintf('WorstSegment_%d_Dashboard', segID), config));
    plotNotes(end + 1) = sprintf('Worst-segment dashboard %d combines speed, power, gear, grade, and braking context for segment-level RCA.', segID);
    close(fig);
end

plotResults = struct('Files', plotFiles, 'Notes', plotNotes);
end

function energyData = localComputeVehicleEnergyFlow(derived, t)
energyData = struct();
energyData.Discharge_kWh = RCA_TrapzFinite(t, max(derived.batteryPower_kW, 0)) / 3600;
energyData.Regen_kWh = RCA_TrapzFinite(t, max(-derived.batteryPower_kW, 0)) / 3600;
energyData.NetBattery_kWh = energyData.Discharge_kWh - energyData.Regen_kWh;
energyData.Aux_kWh = RCA_TrapzFinite(t, max(derived.auxiliaryPower_kW, 0)) / 3600;
energyData.HPR_kWh = RCA_TrapzFinite(t, max(derived.highPowerResistorPower_kW, 0)) / 3600;
energyData.BattLoss_kWh = RCA_TrapzFinite(t, max(derived.batteryLossPower_kW, 0)) / 3600;
energyData.MotorLoss_kWh = RCA_TrapzFinite(t, max(derived.motorLossPower_kW, 0)) / 3600;
energyData.GbxLoss_kWh = RCA_TrapzFinite(t, max(derived.gearboxLossPower_kW, 0)) / 3600;
energyData.Traction_kWh = RCA_TrapzFinite(t, max(derived.tractionPower_kW, 0)) / 3600;
energyData.Friction_kWh = RCA_TrapzFinite(t, max(derived.frictionBrakePower_kW, 0)) / 3600;
energyData.BrakeSplit_kWh = energyData.Regen_kWh + energyData.Friction_kWh;
energyData.NetBus_kWh = max(energyData.Discharge_kWh - energyData.Aux_kWh - energyData.HPR_kWh - energyData.BattLoss_kWh, 0);
if energyData.Discharge_kWh > 0
    energyData.BatteryToWheelEff_pct = 100 * energyData.Traction_kWh / energyData.Discharge_kWh;
else
    energyData.BatteryToWheelEff_pct = NaN;
end
if energyData.BrakeSplit_kWh > 0
    energyData.RegenRecovery_pct = 100 * energyData.Regen_kWh / energyData.BrakeSplit_kWh;
else
    energyData.RegenRecovery_pct = NaN;
end
end

function localDrawEnergyNode(ax, position, labelText, faceColor, fontSize)
if nargin < 5
    fontSize = 10;
end
rectangle(ax, 'Position', position, 'Curvature', 0.04, 'FaceColor', localLightenColor(faceColor, 0.75), ...
    'EdgeColor', localLightenColor(faceColor, 0.25), 'LineWidth', 1.4);
text(ax, position(1) + position(3) / 2, position(2) + position(4) / 2, labelText, ...
    'HorizontalAlignment', 'center', 'VerticalAlignment', 'middle', 'FontWeight', 'bold', 'FontSize', fontSize, 'Interpreter', 'none');
end

function localDrawEnergyArrow(ax, startPoint, endPoint, labelText, labelOffset)
if nargin < 5
    labelOffset = [0.00 0.03];
end
dx = endPoint(1) - startPoint(1);
dy = endPoint(2) - startPoint(2);
quiver(ax, startPoint(1), startPoint(2), dx, dy, 0, 'Color', [0.15 0.15 0.15], ...
    'LineWidth', 1.5, 'MaxHeadSize', 0.7);
text(ax, startPoint(1) + 0.5 * dx + labelOffset(1), startPoint(2) + 0.5 * dy + labelOffset(2), labelText, ...
    'HorizontalAlignment', 'center', 'VerticalAlignment', 'middle', 'FontSize', 9, ...
    'BackgroundColor', 'w', 'Margin', 0.8, 'Interpreter', 'none');
end

function outColor = localLightenColor(inColor, factor)
inColor = double(inColor(:)');
outColor = inColor + (1 - inColor) .* factor;
outColor = min(max(outColor, 0), 1);
end

function [causeNames, contributionSums] = localAggregateCauseContributions(rootCauseRanking)
causeNames = unique(rootCauseRanking.CauseName, 'stable');
contributionSums = zeros(numel(causeNames), 1);
for iCause = 1:numel(causeNames)
    mask = rootCauseRanking.CauseName == causeNames(iCause);
    contributionSums(iCause) = sum(rootCauseRanking.Contribution_pct(mask), 'omitnan');
end
[contributionSums, order] = sort(contributionSums, 'descend');
causeNames = causeNames(order);
end
