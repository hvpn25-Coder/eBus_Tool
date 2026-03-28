function plotResults = RCA_GenerateVehiclePlots(analysisData, outputPaths, config)
% RCA_GenerateVehiclePlots  Create vehicle-level engineering plots.

derived = analysisData.Derived;
t = derived.time_s;
plotFiles = strings(0, 1);
plotNotes = strings(0, 1);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, derived.vehVel_kmh, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
if ~all(isnan(derived.speedDemand_kmh))
    plot(t, derived.speedDemand_kmh, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
    legend({'Vehicle speed', 'Desired speed'}, 'Location', 'best');
end
title('Vehicle Speed Tracking');
ylabel('Speed (km/h)');
grid on;

subplot(2, 1, 2);
plot(t, derived.vehicleAcceleration_mps2, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.roadSlope_pct, '--', 'Color', config.Plot.Colors.Slope, 'LineWidth', config.Plot.LineWidth);
title('Acceleration and Road Slope');
xlabel('Time (s)');
ylabel('Acc (m/s^2) / Slope (%)');
legend({'Acceleration', 'Road slope'}, 'Location', 'best');
grid on;
plotFiles(end + 1) = string(RCA_SaveFigure(fig, outputPaths.FiguresVehicle, 'Vehicle_Speed_Tracking', config));
plotNotes(end + 1) = "Speed tracking plot compares demanded and delivered vehicle response and contextualizes response with acceleration and slope.";
close(fig);

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(3, 1, 1);
plot(t, derived.batteryPower_kW, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, derived.auxiliaryPower_kW, '--', 'Color', config.Plot.Colors.Auxiliary, 'LineWidth', config.Plot.LineWidth);
plot(t, derived.motorLossPower_kW + derived.gearboxLossPower_kW + derived.batteryLossPower_kW, ':', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
plot(t, zeros(size(t)), '-', 'Color', config.Plot.Colors.Neutral, 'LineWidth', 0.8);
title('Vehicle Power Overview (Battery Discharge +, Charge -)');
ylabel('Power (kW)');
legend({'Battery power (RCA sign)', 'Auxiliary power', 'Logged loss power', 'Zero line'}, 'Location', 'best');
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
    ylabel('Share of battery discharge (%)');
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
