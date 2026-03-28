function result = Analyze_PowerTrainController(analysisData, outputPaths, config)
% Analyze_PowerTrainController  Torque demand and saturation analysis.

result = localInitResult("POWER TRAIN CONTROLLER", {'emot1_dem_trq', 'emot2_dem_trq'}, ...
    {'emot1_act_trq', 'emot2_act_trq', 'emot1_max_av_trq', 'emot2_max_av_trq'});
t = analysisData.Derived.time_s;
demand = analysisData.Derived.torqueDemandTotal_Nm;
actual = analysisData.Derived.torqueActualTotal_Nm;
limitPos = analysisData.Derived.torquePositiveLimit_Nm;

rows = cell(0, 7);
summary = strings(0, 1);
if all(isnan(demand))
    result.Warnings(end + 1) = "Controller torque demand signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Power Train Controller", strings(0, 1), strings(0, 1));
    return;
end

torqueErr = demand - actual;
rows = RCA_AddKPI(rows, 'Torque Demand Mean', mean(demand, 'omitnan'), 'Nm', 'Operation', 'Power Train Controller', 'emot1_dem_trq + emot2_dem_trq', 'Complete if both demand signals are available.');
rows = RCA_AddKPI(rows, 'Torque Tracking MAE', mean(abs(torqueErr), 'omitnan'), 'Nm', 'Performance', 'Power Train Controller', 'demand and actual torque', 'Measures delivered torque fidelity.');
rows = RCA_AddKPI(rows, 'Positive Torque Shortfall 95th Percentile', prctile(max(torqueErr, 0), 95), 'Nm', 'Performance', 'Power Train Controller', 'demand and actual torque', 'High values indicate under-delivery during propulsion.');

if ~all(isnan(limitPos))
    limitUse = mean(demand > config.Thresholds.LimitUsageFraction .* limitPos, 'omitnan') * 100;
    rows = RCA_AddKPI(rows, 'Command Near Positive Torque Limit', limitUse, '%', 'Performance', 'Power Train Controller', 'demand torque + max available torque', 'Near-limit operation indicates the controller is requesting all available propulsion.');
    summary(end + 1) = sprintf('Controller saturation evidence: %.1f%% of samples request at least %.0f%% of positive torque capability.', ...
        limitUse, config.Thresholds.LimitUsageFraction * 100);
end

recs = strings(0, 1);
evidence = strings(0, 1);
if mean(max(torqueErr, 0), 'omitnan') > 50
    recs(end + 1) = "Review torque arbitration and limiter coordination; persistent positive torque shortfall points to demand clipping or delayed torque release.";
    evidence(end + 1) = sprintf('Average positive torque shortfall is %.1f Nm.', mean(max(torqueErr, 0), 'omitnan'));
end
if ~all(isnan(limitPos)) && mean(demand > config.Thresholds.LimitUsageFraction .* limitPos, 'omitnan') > 0.05
    recs(end + 1) = "Distinguish controller saturation from plant limitation in reports; the controller frequently requests the positive torque ceiling.";
    evidence(end + 1) = "Demand operates near the available torque envelope for more than 5% of samples.";
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, demand, 'Color', config.Plot.Colors.Demand, 'LineWidth', config.Plot.LineWidth); hold on;
plot(t, actual, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth);
if ~all(isnan(limitPos))
    plot(t, limitPos, '--', 'Color', config.Plot.Colors.Warning, 'LineWidth', config.Plot.LineWidth);
    legend({'Demand torque', 'Actual torque', 'Positive torque limit'}, 'Location', 'best');
else
    legend({'Demand torque', 'Actual torque'}, 'Location', 'best');
end
title('Power Train Controller Torque Demand and Delivery');
ylabel('Torque (Nm)');
grid on;

subplot(2, 1, 2);
plot(t, torqueErr, 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
title('Torque Tracking Error');
xlabel('Time (s)');
ylabel('Demand - Actual (Nm)');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'PowerTrainController'), 'PowerTrainController_Torque', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Power Train Controller", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
