package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Machine-readable ranked task matches for one English keyword query. */
public record TaskKeywordMatchReport(
    GridGrindProtocolVersion protocolVersion,
    String query,
    List<String> normalizedTerms,
    List<String> unmatchedTerms,
    List<String> suggestedIntentTags,
    List<Candidate> candidates) {
  public TaskKeywordMatchReport {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    query = CliDiscoveryValidation.requireNonBlank(query, "query");
    normalizedTerms = CliDiscoveryValidation.copyStrings(normalizedTerms, "normalizedTerms");
    unmatchedTerms = CliDiscoveryValidation.copyStrings(unmatchedTerms, "unmatchedTerms");
    suggestedIntentTags =
        CliDiscoveryValidation.copyStrings(suggestedIntentTags, "suggestedIntentTags");
    Objects.requireNonNull(candidates, "candidates must not be null");
    List<Candidate> copy = new java.util.ArrayList<>(candidates.size());
    Set<String> taskIds = new LinkedHashSet<>();
    for (Candidate candidate : candidates) {
      Candidate value = Objects.requireNonNull(candidate, "candidates must not contain nulls");
      if (!taskIds.add(value.taskId())) {
        throw new IllegalArgumentException(
            "candidates must not contain duplicate task ids: " + value.taskId());
      }
      copy.add(value);
    }
    candidates = List.copyOf(copy);
  }

  /** One compact scored task match that points the caller at the next discovery command. */
  public record Candidate(
      String taskId,
      String summary,
      int score,
      List<String> matchedTerms,
      List<String> matchSources) {
    public Candidate {
      taskId = CliDiscoveryValidation.requireNonBlank(taskId, "taskId");
      summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
      if (score <= 0) {
        throw new IllegalArgumentException("score must be positive");
      }
      matchedTerms = CliDiscoveryValidation.copyStrings(matchedTerms, "matchedTerms");
      matchSources = CliDiscoveryValidation.copyStrings(matchSources, "matchSources");
      if (matchedTerms.isEmpty()) {
        throw new IllegalArgumentException("matchedTerms must not be empty");
      }
      if (matchSources.isEmpty()) {
        throw new IllegalArgumentException("matchSources must not be empty");
      }
    }
  }
}
