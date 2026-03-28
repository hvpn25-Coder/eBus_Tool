function result = Analyze_FinalDrive(analysisData, outputPaths, config)
% Analyze_FinalDrive  Final drive and traction delivery checks.

result = localInitResult("FINAL DRIVE", {'net_trac_trq'}, {'whl_force', 'gr_ratio'});
t = analysisData.Derived.time_s;
tractionForce = analysisData.Derived.tractionForce_N;
wheelForce = analysisData.Derived.wheelForce_N;
tractionPower = analysisData.Derived.tractionPower_kW;

rows = cell(0, 7);
summary = strings(0, 1);

if all(isnan(tractionForce)) && all(isnan(wheelForce))
    result.Warnings(end + 1) = "Net traction or wheel force signals are unavailable.";
    result.KPITable = RCA_FinalizeKPITable(rows);
    result.Suggestions = RCA_MakeSuggestionTable("Final Drive", strings(0, 1), strings(0, 1));
    return;
end

rows = RCA_AddKPI(rows, 'Peak Traction Force', max(tractionForce, [], 'omitnan'), 'N', 'Performance', 'Final Drive', 'net_trac_trq or wheel force basis', 'Traction force is the delivered road force proxy.');
rows = RCA_AddKPI(rows, 'Average Positive Traction Power', mean(max(tractionPower, 0), 'omitnan'), 'kW', 'Performance', 'Final Drive', 'traction force + vehicle speed', 'Positive traction power reflects wheel-end propulsion output.');
if ~all(isnan(wheelForce))
    mismatch = mean(abs(tractionForce - wheelForce), 'omitnan');
    rows = RCA_AddKPI(rows, 'Traction to Wheel Force Mismatch', mismatch, 'N', 'Operation', 'Final Drive', 'net_trac_trq + whl_force', 'Large mismatch can indicate scaling or model plumbing issues.');
    summary(end + 1) = sprintf('Final drive force consistency check: mean traction-to-wheel mismatch is %.1f N.', mismatch);
end

recs = strings(0, 1);
evidence = strings(0, 1);
if ~all(isnan(wheelForce)) && mean(abs(tractionForce - wheelForce), 'omitnan') > 500
    recs(end + 1) = "Check final-drive scaling, driveline sign convention, and brake subtraction if traction and wheel forces diverge materially.";
    evidence(end + 1) = "Mean traction-to-wheel force mismatch exceeds 500 N.";
end

fig = figure('Color', 'w', 'Position', config.Plot.FigurePosition);
subplot(2, 1, 1);
plot(t, tractionForce, 'Color', config.Plot.Colors.Vehicle, 'LineWidth', config.Plot.LineWidth); hold on;
if ~all(isnan(wheelForce))
    plot(t, wheelForce, '--', 'Color', config.Plot.Colors.Motor, 'LineWidth', config.Plot.LineWidth);
    legend({'Traction force', 'Wheel force'}, 'Location', 'best');
end
title('Final Drive Force Delivery');
ylabel('Force (N)');
grid on;

subplot(2, 1, 2);
plot(t, tractionPower, 'Color', config.Plot.Colors.Battery, 'LineWidth', config.Plot.LineWidth);
title('Wheel-End Traction Power');
xlabel('Time (s)');
ylabel('Power (kW)');
grid on;

result.FigureFiles = string(RCA_SaveFigure(fig, fullfile(outputPaths.FiguresSubsystem, 'FinalDrive'), 'FinalDrive_Overview', config));
close(fig);

result.Available = true;
result.KPITable = RCA_FinalizeKPITable(rows);
result.SummaryText = summary;
result.Suggestions = RCA_MakeSuggestionTable("Final Drive", recs, evidence);
end

function result = localInitResult(name, requiredSignals, optionalSignals)
result = struct('Name', string(name), 'Available', false, ...
    'RequiredSignals', {requiredSignals}, 'OptionalSignals', {optionalSignals}, ...
    'KPITable', RCA_FinalizeKPITable([]), 'FigureFiles', strings(0, 1), ...
    'SummaryText', strings(0, 1), 'Warnings', strings(0, 1), ...
    'Suggestions', RCA_MakeSuggestionTable(name, strings(0, 1), strings(0, 1)));
end
