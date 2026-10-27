package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Machine-readable ranked recipe matches for one English keyword query. */
public record RecipeKeywordMatchReport(
    ProtocolVersion protocolVersion,
    String query,
    List<String> normalizedTerms,
    List<String> unmatchedTerms,
    List<String> suggestedIntentTags,
    List<Candidate> candidates) {
  public RecipeKeywordMatchReport {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    query = DiscoveryValidation.requireNonBlank(query, "query");
    normalizedTerms = DiscoveryValidation.copyStrings(normalizedTerms, "normalizedTerms");
    unmatchedTerms = DiscoveryValidation.copyStrings(unmatchedTerms, "unmatchedTerms");
    suggestedIntentTags =
        DiscoveryValidation.copyStringsAllowEmpty(suggestedIntentTags, "suggestedIntentTags");
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
      RecipeType type,
      String summary,
      int score,
      List<String> matchedTerms,
      List<String> matchSources) {
    public Candidate {
      recipeId = DiscoveryValidation.requireNonBlank(recipeId, "recipeId");
      Objects.requireNonNull(type, "type must not be null");
      summary = DiscoveryValidation.requireNonBlank(summary, "summary");
      if (score <= 0) {
        throw new IllegalArgumentException("score must be positive");
      }
      matchedTerms = DiscoveryValidation.copyStrings(matchedTerms, "matchedTerms");
      matchSources = DiscoveryValidation.copyStrings(matchSources, "matchSources");
      if (matchedTerms.isEmpty()) {
        throw new IllegalArgumentException("matchedTerms must not be empty");
      }
      if (matchSources.isEmpty()) {
        throw new IllegalArgumentException("matchSources must not be empty");
      }
    }
  }
}
