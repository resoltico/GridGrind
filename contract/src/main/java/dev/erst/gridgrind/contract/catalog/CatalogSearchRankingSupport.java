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
    return firstPresent(
        rankForTokens(lookupId, tokens, 1),
        rankForTokens(qualifiedId, tokens, 1),
        rankForTokens(summary, tokens, 2),
        rankForTokens(searchableText, tokens, 3));
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
    return firstPresent(
        rankForTokens(lookupId, tokens, 4),
        rankForTokens(qualifiedId, tokens, 4),
        rankForTokens(summary, tokens, 5),
        rankForTokens(searchableText, tokens, 5));
  }

  private static Optional<Integer> fallbackRank(
      List<String> tokens,
      String lookupId,
      String qualifiedId,
      String summary,
      String searchableText,
      String combined) {
    return firstPresent(
        rankForTokens(lookupId, tokens, 6),
        rankForTokens(qualifiedId, tokens, 6),
        rankForTokens(summary, tokens, 7),
        rankForTokens(searchableText, tokens, 7),
        rankForTokens(combined, tokens, 8));
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

  private static Optional<Integer> rankForTokens(String haystack, List<String> tokens, int tier) {
    int matches = matchedTokenCount(haystack, tokens);
    if (matches == 0) {
      return Optional.empty();
    }
    int missing = tokens.size() - matches;
    return Optional.of(missing * 1_000 + tier * 10 - matches);
  }

  static int matchedTokenCount(String haystack, List<String> tokens) {
    int matches = 0;
    for (String token : tokens) {
      if (haystack.contains(token)) {
        matches++;
      }
    }
    return matches;
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
