package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Deterministic English keyword matcher layered on top of the published recipe registry. */
public final class RecipeMatcher {
  private RecipeMatcher() {}

  /** Returns ranked recipe matches for one English keyword query string. */
  public static RecipeKeywordMatchReport reportFor(String query) {
    String requestedQuery = Objects.requireNonNull(query, "query must not be null");
    List<String> normalizedTerms = RecipeTerms.normalizedTerms(requestedQuery);
    if (normalizedTerms.isEmpty()) {
      throw new IllegalArgumentException(
          "query must contain at least one searchable term after normalization");
    }
    List<RecipeKeywordMatchReport.Candidate> candidates =
        RecipeRegistry.recipes().stream()
            .map(recipe -> RecipeScorer.candidateFor(recipe, normalizedTerms))
            .flatMap(Optional::stream)
            .sorted(candidateOrdering())
            .toList();
    return new RecipeKeywordMatchReport(
        ProtocolVersion.CURRENT,
        requestedQuery,
        normalizedTerms,
        unmatchedTerms(normalizedTerms, candidates),
        suggestedIntentTags(candidates),
        candidates);
  }

  static Comparator<RecipeKeywordMatchReport.Candidate> candidateOrdering() {
    return Comparator.comparingInt(RecipeKeywordMatchReport.Candidate::score)
        .reversed()
        .thenComparing(candidate -> candidate.matchedTerms().size(), Comparator.reverseOrder())
        .thenComparing(RecipeKeywordMatchReport.Candidate::recipeId);
  }

  static List<String> suggestedIntentTags(List<RecipeKeywordMatchReport.Candidate> candidates) {
    if (candidates.isEmpty()) {
      return RecipeRegistry.intentTags();
    }
    int topScore = candidates.getFirst().score();
    List<TagSuggestion> suggestionScores = new ArrayList<>();
    for (RecipeKeywordMatchReport.Candidate candidate : candidates) {
      if (candidate.score() * 4 < topScore) {
        continue;
      }
      RecipeRegistry.recipeFor(candidate.recipeId())
          .orElseThrow()
          .intentTags()
          .forEach(tag -> mergeTagSuggestion(suggestionScores, tag, candidate.score()));
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

  private static List<String> unmatchedTerms(
      List<String> queryTerms, List<RecipeKeywordMatchReport.Candidate> candidates) {
    Set<String> matched = new LinkedHashSet<>();
    for (RecipeKeywordMatchReport.Candidate candidate : candidates) {
      matched.addAll(candidate.matchedTerms());
    }
    return queryTerms.stream().filter(term -> !matched.contains(term)).toList();
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
