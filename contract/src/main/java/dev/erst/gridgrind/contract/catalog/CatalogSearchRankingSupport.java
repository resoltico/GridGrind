package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Optional;

/** Search rank scoring, top-level group classification, and step-template resolution. */
final class CatalogSearchRankingSupport {
  private CatalogSearchRankingSupport() {}

  static Optional<Integer> searchRank(
      boolean topLevelEntryMatch,
      boolean supportingEntryMatch,
      String normalizedQuery,
      List<String> tokens,
      String lookupId,
      String qualifiedId,
      String summary,
      String searchableText,
      String combined) {
    return firstPresent(
        exactQueryRank(normalizedQuery, lookupId, qualifiedId),
        topLevelRank(topLevelEntryMatch, tokens, lookupId, qualifiedId, summary, searchableText),
        supportingRank(
            supportingEntryMatch, tokens, lookupId, qualifiedId, summary, searchableText),
        fallbackRank(tokens, lookupId, qualifiedId, summary, searchableText, combined));
  }

  private static Optional<Integer> exactQueryRank(
      String normalizedQuery, String lookupId, String qualifiedId) {
    return lookupId.equals(normalizedQuery) || qualifiedId.equals(normalizedQuery)
        ? Optional.of(0)
        : Optional.empty();
  }

  private static Optional<Integer> topLevelRank(
      boolean topLevelEntryMatch,
      List<String> tokens,
      String lookupId,
      String qualifiedId,
      String summary,
      String searchableText) {
    if (!topLevelEntryMatch) {
      return Optional.empty();
    }
    if (containsAllTokens(lookupId, tokens) || containsAllTokens(qualifiedId, tokens)) {
      return Optional.of(1);
    }
    if (containsAllTokens(summary, tokens)) {
      return Optional.of(2);
    }
    return containsAllTokens(searchableText, tokens) ? Optional.of(3) : Optional.empty();
  }

  private static Optional<Integer> supportingRank(
      boolean supportingEntryMatch,
      List<String> tokens,
      String lookupId,
      String qualifiedId,
      String summary,
      String searchableText) {
    if (!supportingEntryMatch) {
      return Optional.empty();
    }
    if (containsAllTokens(lookupId, tokens) || containsAllTokens(qualifiedId, tokens)) {
      return Optional.of(4);
    }
    return containsAllTokens(summary, tokens) || containsAllTokens(searchableText, tokens)
        ? Optional.of(5)
        : Optional.empty();
  }

  private static Optional<Integer> fallbackRank(
      List<String> tokens,
      String lookupId,
      String qualifiedId,
      String summary,
      String searchableText,
      String combined) {
    if (containsAllTokens(lookupId, tokens) || containsAllTokens(qualifiedId, tokens)) {
      return Optional.of(6);
    }
    if (containsAllTokens(summary, tokens) || containsAllTokens(searchableText, tokens)) {
      return Optional.of(7);
    }
    return containsAllTokens(combined, tokens) ? Optional.of(8) : Optional.empty();
  }

  @SafeVarargs
  static <T> Optional<T> firstPresent(Optional<T>... candidates) {
    for (Optional<T> candidate : candidates) {
      if (candidate.isPresent()) {
        return candidate;
      }
    }
    return Optional.empty();
  }

  static boolean containsAllTokens(String haystack, List<String> tokens) {
    for (String token : tokens) {
      if (!haystack.contains(token)) {
        return false;
      }
    }
    return true;
  }

  static boolean isTopLevelPublishedGroup(String catalogGroup) {
    return switch (catalogGroup) {
      case "sourceTypes",
          "persistenceTypes",
          "stepTypes",
          "mutationActionTypes",
          "assertionTypes",
          "inspectionQueryTypes" ->
          true;
      default -> false;
    };
  }

  static Optional<ProtocolStepTemplate> stepTemplateFor(
      CatalogLookupRef ref, CatalogLookupValue lookupValue) {
    if (lookupValue instanceof CatalogLookupValue.EntryLookupValue entryLookupValue
        && isTopLevelPublishedGroup(ref.catalogGroup())) {
      return entryLookupValue.entry().stepTemplate();
    }
    return Optional.empty();
  }
}
