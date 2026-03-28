function result = Analyze_Driver(analysisData, outputPaths, config)
% Analyze_Driver  Driver demand and behaviour analysis.

result = localInitResult("DRIVER", {'acc_pdl', 'brk_pdl'}, {'veh_des_vel', 'veh_vel'});
t = analysisData.Derived.time_s;
movingMask = analysisData.Derived.vehVel_kmh > config.Thresholds.StopSpeed_kmh;

accPedal = RCA_GetSignalData(analysisData.Signals, 'acc_pdl');
brkPedal = RCA_GetSignalData(analysisData.Signals, 'brk_pdl');
desiredSpeed = RCA_GetSignalData(analysisData.Signals, 'veh_des_vel');
rows = cell(0, 7);
summary = strings(0, 1);

if accPedal.Available
    rows = RCA_AddKPI(rows, 'Mean Accelerator Pedal', mean(analysisData.Derived.accPedal_pct, 'omitnan'), '%', 'Operation', 'Driver', 'acc_pdl', 'Available from driver demand.');
    rows = RCA_AddKPI(rows, 'Accelerator Pedal 95th Percentile', prctile(analysisData.Derived.accPedal_pct, 95), '%', 'Performance', 'Driver', 'acc_pdl', 'Reflects aggressive drive demand level.');
end

if brkPedal.Available
    rows = RCA_AddKPI(rows, 'Mean Brake Pedal', mean(analysisData.Derived.brkPedal_pct, 'omitnan'), '%', 'Operation', 'Driver', 'brk_pdl', 'Available from driver demand.');
    rows = RCA_AddKPI(rows, 'Brake Pedal Active Time Share', 100 * mean(analysisData.Derived.brkPedal_pct > 1, 'omitnan'), '%', 'Operation', 'Driver', 'brk_pdl', '1% pedal threshold is heuristic and editable.');
end

if accPedal.Available && brkPedal.Available
    overlap = mean((analysisData.Derived.accPedal_pct > 5) & (analysisData.Derived.brkPedal_pct > 5) & movingMask, 'omitnan') * 100;
    coasting = mean((analysisData.Derived.accPedal_pct < 1) & (analysisData.Derived.brkPedal_pct < 1) & movingMask, 'omitnan') * 100;
    rows = RCA_AddKPI(rows, 'Pedal Overlap Time Share', overlap, '%', 'Operation', 'Driver', 'acc_pdl + brk_pdl', 'Pedal overlap is a heuristic quality indicator.');
    rows = RCA_AddKPI(rows, 'Coasting Time Share', coasting, '%', 'Efficiency', 'Driver', 'acc_pdl + brk_pdl + veh_vel', 'Coasting can support efficient operation when route allows.');
    summary(end + 1) = sprintf('Driver demand quality: pedal overlap is %.2f%% of moving time and coasting share is %.1f%%.', overlap, coasting);
end

if desiredSpeed.Available
    trackingErr = abs(analysisData.Derived.speedDemand_kmh - analysisData.Derived.vehVel_kmh);
    rows = RCA_AddKPI(rows, 'Driver-to-Vehicle Tracking MAE', mean(trackingErr, 'omitnan'), 'km/h', 'Performance', 'Driver', 'veh_des_vel + veh_vel', 'Measures how the vehicle follows the requested cycle.');
end

recs = strings(0, 1);
evidence = strings(0, 1);
if brkPedal.Available && accPedal.Available
    overlap = mean((analysisData.Derived.accPedal_pct > 5) & (analysisData.Derived.brkPedal_pct > 5) & movingMask, 'omitnan') * 100;
    if overlap > 1
        recs(end + 1) = "Review driver model transitions around braking and re-acceleration; simultaneous pedal activity can hide controller or brake-blending behaviour.";
        evidence(end + 1) = sprintf('Pedal overlap reached %.2f%% of moving time.', overlap);
    end
end
if accPedal.Available && mean(analysisData.Derived.accPedal_pct > 80, 'omitnan') > 0.10
    recs(end + 1) = "A large fraction of full-throttle requests suggests a demanding cycle; keep this separate from subsystem inefficiency when assigning blame.";
    evidence(end + 1) = "Accelerator pedal spends more than 10% of time above 80%.";
end

if accPedal.Available || brkPedal.Available
    fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
    subplot(2, 1, 1);
    legendEntries = {};
    if accPedal.Available
        plot(t, analysisData.Derived.accPedal_pct, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
        legendEntries{end + 1} = 'Accelerator';
    end
    if brkPedal.Available
        plot(t, analysisData.Derived.brkPedal_pct, 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
        legendEntries{end + 1} = 'Brake';
    end
    title('Driver Pedal Activity');
    ylabel('Pedal (%)');
    if ~isempty(legendEntries)
        legend(legendEntries, 'Location', 'best');
    end
    grid on;

    subplot(2, 1, 2);
    plot(t, analysisData.Derived.vehVel_kmh, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
    if desiredSpeed.Available
        plot(t, analysisData.Derived.speedDemand_kmh, '--', 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth);
        legend({'Vehicle speed', 'Desired speed'}, 'Location', 'best');
    end
    xlabel('Time (s)');
    ylabel('Speed (km/h)');
    title('Driver Demand Versus Vehicle Response');
    grid on;
    result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'Driver'), 'Driver_Demand_Overview', config));
    close(fig);
end

result.Available = ~isempty(rows);
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Driver", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
