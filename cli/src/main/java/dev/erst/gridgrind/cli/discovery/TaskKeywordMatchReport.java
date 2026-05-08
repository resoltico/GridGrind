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
      if (!taskIds.add(value.task().id())) {
        throw new IllegalArgumentException(
            "candidates must not contain duplicate task ids: " + value.task().id());
      }
      copy.add(value);
    }
    candidates = List.copyOf(copy);
  }

  /** One scored task match plus the starter scaffold derived from the exact task descriptor. */
  public record Candidate(
      TaskEntry task,
      int score,
      List<String> matchedTerms,
      List<String> reasons,
      TaskPlanTemplate starterTemplate) {
    public Candidate {
      Objects.requireNonNull(task, "task must not be null");
      if (score <= 0) {
        throw new IllegalArgumentException("score must be positive");
      }
      matchedTerms = CliDiscoveryValidation.copyStrings(matchedTerms, "matchedTerms");
      reasons = CliDiscoveryValidation.copyStrings(reasons, "reasons");
      Objects.requireNonNull(starterTemplate, "starterTemplate must not be null");
      if (matchedTerms.isEmpty()) {
        throw new IllegalArgumentException("matchedTerms must not be empty");
      }
      if (reasons.isEmpty()) {
        throw new IllegalArgumentException("reasons must not be empty");
      }
      if (!task.id().equals(starterTemplate.task().id())) {
        throw new IllegalArgumentException("starterTemplate task id must match candidate task id");
      }
    }
  }
}
