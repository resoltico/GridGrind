package dev.erst.gridgrind.contract.catalog;

import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Stream;

/** Search match grouping, deduplication, and owner-match resolution for catalog search results. */
final class CatalogSearchAggregationSupport {
  private CatalogSearchAggregationSupport() {}

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static List<RankedSearchMatch> groupSearchMatches(
      Catalog catalog, List<RankedSearchMatch> matches, String normalizedQuery) {
    Set<String> nestedGroupNames =
        catalog.nestedTypes().stream()
            .map(NestedTypeGroup::group)
            .collect(java.util.stream.Collectors.toSet());
    Set<String> plainGroupNames =
        catalog.plainTypes().stream()
            .map(PlainTypeGroup::group)
            .collect(java.util.stream.Collectors.toSet());
    Map<String, SearchAggregate> aggregates = new LinkedHashMap<>();
    Map<String, RankedSearchMatch> standalone = new LinkedHashMap<>();
    for (RankedSearchMatch rankedMatch : matches) {
      CatalogSearchMatch match = rankedMatch.match();
      if (isOperationEntryMatch(match)) {
        aggregateFor(
                aggregates, match.qualifiedId(), rankedMatch.rank(), ownerMatch(catalog, match))
            .includeDirectRank(rankedMatch.rank());
        continue;
      }
      if (isSupportingMatch(match, nestedGroupNames, plainGroupNames)) {
        boolean attached = attachSupportingMatch(catalog, aggregates, rankedMatch, match);
        if (attached && isExactLookupMatch(match, normalizedQuery)) {
          recordStandalone(standalone, rankedMatch.rank(), match);
        }
        if (attached) {
          continue;
        }
        CatalogSearchMatch collapsedMatch =
            isExactLookupMatch(match, normalizedQuery)
                ? match
                : collapsedSupportingEntryMatch(catalog, match, nestedGroupNames, plainGroupNames)
                    .orElse(match);
        recordStandalone(standalone, rankedMatch.rank(), collapsedMatch);
        continue;
      }
      recordStandalone(standalone, rankedMatch.rank(), match);
    }
    return Stream.concat(
            aggregates.values().stream().map(SearchAggregate::toRankedMatch),
            standalone.values().stream())
        .sorted(
            Comparator.comparingInt(RankedSearchMatch::rank)
                .thenComparing(match -> match.match().qualifiedId()))
        .toList();
  }

  private static void recordStandalone(
      Map<String, RankedSearchMatch> standalone, int rank, CatalogSearchMatch match) {
    RankedSearchMatch candidate = new RankedSearchMatch(rank, match);
    standalone.merge(
        match.qualifiedId(),
        candidate,
        (left, right) -> left.rank() <= right.rank() ? left : right);
  }

  private static SearchAggregate aggregateFor(
      Map<String, SearchAggregate> aggregates,
      String qualifiedId,
      int rank,
      CatalogSearchMatch owner) {
    SearchAggregate aggregate = aggregates.get(qualifiedId);
    if (aggregate != null) {
      return aggregate;
    }
    SearchAggregate created = new SearchAggregate(rank, owner);
    aggregates.put(qualifiedId, created);
    return created;
  }

  private static boolean attachSupportingMatch(
      Catalog catalog,
      Map<String, SearchAggregate> aggregates,
      RankedSearchMatch rankedMatch,
      CatalogSearchMatch match) {
    boolean attached = false;
    for (String relatedEntryId : match.relatedEntryIds()) {
      CatalogSearchMatch owner = requireOwnerMatchForQualifiedId(catalog, relatedEntryId);
      attached = true;
      aggregateFor(aggregates, relatedEntryId, rankedMatch.rank() + 1, owner)
          .attachSupportingMatch(rankedMatch.rank() + 1, match);
    }
    return attached;
  }

  private static boolean isOperationEntryMatch(CatalogSearchMatch match) {
    return "ENTRY".equals(match.kind())
        && CatalogSearchRankingSupport.isTopLevelPublishedGroup(match.catalogGroup());
  }

  private static boolean isSupportingMatch(
      CatalogSearchMatch match, Set<String> nestedGroupNames, Set<String> plainGroupNames) {
    if ("nestedTypes".equals(match.catalogGroup()) || "plainTypes".equals(match.catalogGroup())) {
      return true;
    }
    return nestedGroupNames.contains(match.catalogGroup())
        || plainGroupNames.contains(match.catalogGroup());
  }

  private static Optional<CatalogSearchMatch> collapsedSupportingEntryMatch(
      Catalog catalog,
      CatalogSearchMatch match,
      Set<String> nestedGroupNames,
      Set<String> plainGroupNames) {
    if (!"ENTRY".equals(match.kind())) {
      return Optional.empty();
    }
    if (nestedGroupNames.contains(match.catalogGroup())) {
      return groupLookupMatchFor(catalog, "nestedTypes", match.catalogGroup());
    }
    if (plainGroupNames.contains(match.catalogGroup())) {
      return groupLookupMatchFor(catalog, "plainTypes", match.catalogGroup());
    }
    return Optional.empty();
  }

  private static boolean isExactLookupMatch(CatalogSearchMatch match, String normalizedQuery) {
    return match.lookupId().toLowerCase(Locale.ROOT).equals(normalizedQuery)
        || match.qualifiedId().toLowerCase(Locale.ROOT).equals(normalizedQuery);
  }

  private static CatalogSearchMatch ownerMatch(Catalog catalog, CatalogSearchMatch match) {
    return ownerMatchForQualifiedId(catalog, match.qualifiedId()).orElse(match);
  }

  private static CatalogSearchMatch requireOwnerMatchForQualifiedId(
      Catalog catalog, String qualifiedId) {
    return ownerMatchForQualifiedId(catalog, qualifiedId)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Catalog related entry id did not resolve: " + qualifiedId));
  }

  private static Optional<CatalogSearchMatch> ownerMatchForQualifiedId(
      Catalog catalog, String qualifiedId) {
    return CatalogRefResolutionSupport.allLookupRefs(catalog).stream()
        .filter(
            ref ->
                ref.qualifiedId().equals(qualifiedId)
                    && ref.value() instanceof CatalogLookupValue.EntryLookupValue
                    && CatalogSearchRankingSupport.isTopLevelPublishedGroup(ref.catalogGroup()))
        .findFirst()
        .map(
            ref ->
                new CatalogSearchMatch(
                    ref.catalogGroup(),
                    ref.lookupId(),
                    ref.qualifiedId(),
                    ref.value().kind(),
                    ref.value().summary(),
                    CatalogSearchRankingSupport.stepTemplateFor(ref, ref.value()),
                    List.of(),
                    List.of()));
  }

  private static Optional<CatalogSearchMatch> groupLookupMatchFor(
      Catalog catalog, String catalogGroup, String lookupId) {
    return CatalogRefResolutionSupport.allLookupRefs(catalog).stream()
        .filter(ref -> ref.catalogGroup().equals(catalogGroup) && ref.lookupId().equals(lookupId))
        .findFirst()
        .map(
            ref ->
                new CatalogSearchMatch(
                    ref.catalogGroup(),
                    ref.lookupId(),
                    ref.qualifiedId(),
                    ref.value().kind(),
                    ref.value().summary(),
                    CatalogSearchRankingSupport.stepTemplateFor(ref, ref.value()),
                    CatalogShapeTextSupport.relatedEntryIdsFor(catalog, ref, ref.value()),
                    List.of()));
  }

  /** Aggregates grouped search state for one top-level published catalog owner. */
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static final class SearchAggregate {
    private int rank;
    private final CatalogSearchMatch owner;
    private final Map<String, CatalogSearchMatch> supportingMatches = new LinkedHashMap<>();

    private SearchAggregate(int rank, CatalogSearchMatch owner) {
      this.rank = rank;
      this.owner = Objects.requireNonNull(owner, "owner must not be null");
    }

    private void includeDirectRank(int directRank) {
      rank = Math.min(rank, directRank);
    }

    private void attachSupportingMatch(int supportRank, CatalogSearchMatch support) {
      rank = Math.min(rank, supportRank);
      supportingMatches.putIfAbsent(support.qualifiedId(), support);
    }

    private RankedSearchMatch toRankedMatch() {
      return new RankedSearchMatch(
          rank,
          new CatalogSearchMatch(
              owner.catalogGroup(),
              owner.lookupId(),
              owner.qualifiedId(),
              owner.kind(),
              owner.summary(),
              owner.stepTemplate(),
              owner.relatedEntryIds(),
              List.copyOf(supportingMatches.values())));
    }
  }
}
