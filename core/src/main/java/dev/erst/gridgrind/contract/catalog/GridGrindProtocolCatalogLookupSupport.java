package dev.erst.gridgrind.contract.catalog;

import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

/** Lookup and fuzzy-search helpers split out of the protocol catalog registry. */
final class GridGrindProtocolCatalogLookupSupport {
  private GridGrindProtocolCatalogLookupSupport() {}

  static Optional<TypeEntry> entryFor(Catalog catalog, String idOrQualifiedId) {
    List<CatalogRefResolutionSupport.CatalogEntryRef> matches =
        CatalogRefResolutionSupport.matchingEntryRefs(catalog, idOrQualifiedId);
    return matches.size() == 1 ? Optional.of(matches.getFirst().entry()) : Optional.empty();
  }

  static Optional<Object> lookupValueFor(Catalog catalog, String idOrQualifiedId) {
    List<CatalogLookupRef> matches =
        CatalogRefResolutionSupport.matchingLookupRefs(catalog, idOrQualifiedId);
    return matches.size() == 1
        ? Optional.of(CatalogLookupValueSupport.rawValue(matches.getFirst().value()))
        : Optional.empty();
  }

  static List<String> matchingEntryIds(Catalog catalog, String idOrQualifiedId) {
    return CatalogRefResolutionSupport.matchingEntryRefs(catalog, idOrQualifiedId).stream()
        .map(CatalogRefResolutionSupport.CatalogEntryRef::qualifiedId)
        .toList();
  }

  static List<String> matchingLookupIds(Catalog catalog, String idOrQualifiedId) {
    return CatalogRefResolutionSupport.matchingLookupRefs(catalog, idOrQualifiedId).stream()
        .map(CatalogLookupRef::qualifiedId)
        .toList();
  }

  static CatalogSearchResult search(Catalog catalog, String query) {
    String trimmedQuery = CatalogRecordValidation.requireNonBlank(query, "query").trim();
    String normalizedQuery = trimmedQuery.toLowerCase(Locale.ROOT);
    List<String> tokens = List.of(normalizedQuery.split("\\s+"));
    List<RankedSearchMatch> rankedMatches =
        CatalogRefResolutionSupport.allLookupRefs(catalog).stream()
            .map(ref -> searchMatch(catalog, ref, normalizedQuery, tokens))
            .flatMap(Optional::stream)
            .sorted(
                Comparator.comparingInt(RankedSearchMatch::rank)
                    .thenComparing(match -> match.match().qualifiedId()))
            .toList();
    return new CatalogSearchResult(
        catalog.protocolVersion(),
        trimmedQuery,
        CatalogSearchAggregationSupport.groupSearchMatches(catalog, rankedMatches, normalizedQuery)
            .stream()
            .map(RankedSearchMatch::match)
            .toList());
  }

  static List<RankedSearchMatch> groupSearchMatches(
      Catalog catalog, List<RankedSearchMatch> matches, String normalizedQuery) {
    return CatalogSearchAggregationSupport.groupSearchMatches(catalog, matches, normalizedQuery);
  }

  private static Optional<RankedSearchMatch> searchMatch(
      Catalog catalog, CatalogLookupRef ref, String normalizedQuery, List<String> tokens) {
    CatalogLookupValue lookupValue = ref.value();
    boolean topLevelEntryMatch =
        lookupValue instanceof CatalogEntryLookupValue
            && CatalogSearchRankingSupport.isTopLevelPublishedGroup(ref.catalogGroup());
    boolean supportingEntryMatch =
        lookupValue instanceof CatalogEntryLookupValue
            && !CatalogSearchRankingSupport.isTopLevelPublishedGroup(ref.catalogGroup());
    String lookupId = ref.lookupId().toLowerCase(Locale.ROOT);
    String qualifiedId = ref.qualifiedId().toLowerCase(Locale.ROOT);
    String catalogGroup = ref.catalogGroup().toLowerCase(Locale.ROOT);
    String summary = CatalogLookupValueSupport.summary(lookupValue).toLowerCase(Locale.ROOT);
    String searchableText =
        CatalogLookupValueSupport.searchableText(catalog, lookupValue, topLevelEntryMatch)
            .toLowerCase(Locale.ROOT);
    String combined =
        lookupId + " " + qualifiedId + " " + catalogGroup + " " + summary + " " + searchableText;
    Optional<Integer> rank =
        CatalogSearchRankingSupport.searchRank(
            topLevelEntryMatch,
            supportingEntryMatch,
            normalizedQuery,
            tokens,
            lookupId,
            qualifiedId,
            summary,
            searchableText,
            combined);
    if (rank.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(
        new RankedSearchMatch(
            rank.orElseThrow(),
            new CatalogSearchMatch(
                ref.catalogGroup(),
                ref.lookupId(),
                ref.qualifiedId(),
                CatalogLookupValueSupport.kind(lookupValue),
                CatalogLookupValueSupport.summary(lookupValue),
                CatalogSearchRankingSupport.stepTemplateFor(ref, lookupValue),
                CatalogShapeTextSupport.relatedEntryIdsFor(catalog, ref, lookupValue),
                List.of())));
  }
}
