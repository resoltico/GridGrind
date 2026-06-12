package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskKeywordMatchReport;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Deterministic English keyword matcher layered on top of the public task descriptor catalog. */
final class GridGrindTaskKeywordMatcher {
  private GridGrindTaskKeywordMatcher() {}

  /** Returns ranked task matches for one English keyword query string. */
  static TaskKeywordMatchReport reportFor(String query) {
    String requestedQuery = Objects.requireNonNull(query, "query must not be null");
    List<String> normalizedTerms = GridGrindTaskKeywordText.normalizedTerms(requestedQuery);
    if (normalizedTerms.isEmpty()) {
      throw new IllegalArgumentException(
          "query must contain at least one searchable term after normalization");
    }
    List<TaskKeywordMatchReport.Candidate> candidates =
        GridGrindTaskCatalog.catalog().tasks().stream()
            .map(task -> GridGrindTaskKeywordCandidateScorer.candidateFor(task, normalizedTerms))
            .flatMap(Optional::stream)
            .sorted(candidateOrdering())
            .toList();
    return new TaskKeywordMatchReport(
        GridGrindProtocolVersion.current(),
        requestedQuery,
        normalizedTerms,
        unmatchedTerms(normalizedTerms, candidates),
        suggestedIntentTags(normalizedTerms, candidates),
        candidates);
  }

  private static List<String> unmatchedTerms(
      List<String> queryTerms, List<TaskKeywordMatchReport.Candidate> candidates) {
    Set<String> matched = new LinkedHashSet<>();
    for (TaskKeywordMatchReport.Candidate candidate : candidates) {
      matched.addAll(candidate.matchedTerms());
    }
    return queryTerms.stream().filter(term -> !matched.contains(term)).toList();
  }

  static Comparator<TaskKeywordMatchReport.Candidate> candidateOrdering() {
    return Comparator.comparingInt(TaskKeywordMatchReport.Candidate::score)
        .reversed()
        .thenComparing(candidate -> candidate.matchedTerms().size(), Comparator.reverseOrder())
        .thenComparing(TaskKeywordMatchReport.Candidate::taskId);
  }

  static List<String> suggestedIntentTags(
      List<String> queryTerms, List<TaskKeywordMatchReport.Candidate> candidates) {
    if (candidates.isEmpty()) {
      return List.of();
    }
    int topScore = candidates.getFirst().score();
    List<TagSuggestion> suggestionScores = new ArrayList<>();
    for (TaskKeywordMatchReport.Candidate candidate : candidates) {
      if (candidate.score() * 4 < topScore) {
        continue;
      }
      TaskEntry task = GridGrindTaskCatalog.entryFor(candidate.taskId()).orElseThrow();
      for (String tag : task.discoveryProfile().intentTags()) {
        int overlapCount =
            GridGrindTaskKeywordText.intersection(
                    queryTerms, GridGrindTaskKeywordText.normalizedTerms(tag))
                .size();
        int weightedScore = candidate.score() + (overlapCount * 10_000);
        mergeTagSuggestion(suggestionScores, tag, weightedScore);
      }
    }
    return suggestionScores.stream()
        .sorted(
            Comparator.comparingInt(TagSuggestion::score)
                .reversed()
                .thenComparing(TagSuggestion::tag))
        .limit(8)
        .map(TagSuggestion::tag)
        .toList();
  }

  private static void mergeTagSuggestion(
      List<TagSuggestion> suggestions, String tag, int weightedScore) {
    int existingIndex = -1;
    for (int index = 0; index < suggestions.size(); index++) {
      if (suggestions.get(index).tag().equals(tag)) {
        existingIndex = index;
        break;
      }
    }
    if (existingIndex < 0) {
      suggestions.add(new TagSuggestion(tag, weightedScore));
      return;
    }
    TagSuggestion existing = suggestions.get(existingIndex);
    if (weightedScore > existing.score()) {
      suggestions.set(existingIndex, new TagSuggestion(tag, weightedScore));
    }
  }

  private record TagSuggestion(String tag, int score) {}
}
