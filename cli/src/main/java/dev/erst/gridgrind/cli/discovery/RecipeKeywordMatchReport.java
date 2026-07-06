package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Machine-readable ranked recipe matches for one English keyword query. */
public record RecipeKeywordMatchReport(
    GridGrindProtocolVersion protocolVersion,
    String query,
    List<String> normalizedTerms,
    List<String> unmatchedTerms,
    List<String> suggestedIntentTags,
    List<Candidate> candidates) {
  public RecipeKeywordMatchReport {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    query = CliDiscoveryValidation.requireNonBlank(query, "query");
    normalizedTerms = CliDiscoveryValidation.copyStrings(normalizedTerms, "normalizedTerms");
    unmatchedTerms = CliDiscoveryValidation.copyStrings(unmatchedTerms, "unmatchedTerms");
    suggestedIntentTags =
        CliDiscoveryValidation.copyStringsAllowEmpty(suggestedIntentTags, "suggestedIntentTags");
    Objects.requireNonNull(candidates, "candidates must not be null");
    List<Candidate> copy = new java.util.ArrayList<>(candidates.size());
    Set<String> recipeIds = new LinkedHashSet<>();
    for (Candidate candidate : candidates) {
      Candidate value = Objects.requireNonNull(candidate, "candidates must not contain nulls");
      if (!recipeIds.add(value.recipeId())) {
        throw new IllegalArgumentException(
            "candidates must not contain duplicate recipe ids: " + value.recipeId());
      }
      copy.add(value);
    }
    candidates = List.copyOf(copy);
  }

  /** One compact scored recipe match that points the caller at the next discovery command. */
  public record Candidate(
      String recipeId,
      RecipeView view,
      String summary,
      int score,
      List<String> matchedTerms,
      List<String> matchSources) {
    public Candidate {
      recipeId = CliDiscoveryValidation.requireNonBlank(recipeId, "recipeId");
      Objects.requireNonNull(view, "view must not be null");
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
