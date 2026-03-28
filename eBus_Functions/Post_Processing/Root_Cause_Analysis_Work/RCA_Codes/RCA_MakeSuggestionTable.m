function suggestionTable = RCA_MakeSuggestionTable(subsystem, recommendations, evidence)
% RCA_MakeSuggestionTable  Create a standard subsystem suggestion table.

if nargin < 2 || isempty(recommendations)
    suggestionTable = table(strings(0, 1), strings(0, 1), strings(0, 1), ...
        'VariableNames', {'Subsystem', 'Recommendation', 'Evidence'});
    return;
end

recommendations = string(recommendations(:));
if nargin < 3 || isempty(evidence)
    evidence = repmat("", size(recommendations));
else
    evidence = string(evidence(:));
    if numel(evidence) == 1 && numel(recommendations) > 1
        evidence = repmat(evidence, size(recommendations));
    end
end

subsystem = repmat(string(subsystem), size(recommendations));
suggestionTable = table(subsystem, recommendations, evidence, ...
    'VariableNames', {'Subsystem', 'Recommendation', 'Evidence'});
end
