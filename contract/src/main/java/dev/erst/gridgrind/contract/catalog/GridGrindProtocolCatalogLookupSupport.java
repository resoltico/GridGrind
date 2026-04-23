package dev.erst.gridgrind.contract.catalog;

import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Function;
import java.util.stream.Stream;

/** Lookup and fuzzy-search helpers split out of the protocol catalog registry. */
final class GridGrindProtocolCatalogLookupSupport {
  private GridGrindProtocolCatalogLookupSupport() {}

  static Optional<TypeEntry> entryFor(Catalog catalog, String idOrQualifiedId) {
    List<CatalogEntryRef> matches = matchingEntryRefs(catalog, idOrQualifiedId);
    return matches.size() == 1 ? Optional.of(matches.getFirst().entry()) : Optional.empty();
  }

  static Optional<Object> lookupValueFor(Catalog catalog, String idOrQualifiedId) {
    List<CatalogLookupRef> matches = matchingLookupRefs(catalog, idOrQualifiedId);
    return matches.size() == 1 ? Optional.of(matches.getFirst().value()) : Optional.empty();
  }

  static List<String> matchingEntryIds(Catalog catalog, String idOrQualifiedId) {
    return matchingEntryRefs(catalog, idOrQualifiedId).stream()
        .map(CatalogEntryRef::qualifiedId)
        .toList();
  }

  static List<String> matchingLookupIds(Catalog catalog, String idOrQualifiedId) {
    return matchingLookupRefs(catalog, idOrQualifiedId).stream()
        .map(CatalogLookupRef::qualifiedId)
        .toList();
  }

  static CatalogSearchResult search(Catalog catalog, String query) {
    String trimmedQuery = CatalogRecordValidation.requireNonBlank(query, "query").trim();
    String normalizedQuery = trimmedQuery.toLowerCase(Locale.ROOT);
    List<String> tokens = List.of(normalizedQuery.split("\\s+"));
    List<CatalogSearchMatch> matches =
        allLookupRefs(catalog).stream()
            .map(ref -> searchMatch(ref, normalizedQuery, tokens))
            .filter(Objects::nonNull)
            .sorted(
                Comparator.comparingInt(RankedSearchMatch::rank)
                    .thenComparing(match -> match.match().qualifiedId()))
            .map(RankedSearchMatch::match)
            .toList();
    return new CatalogSearchResult(trimmedQuery, matches);
  }

  private static RankedSearchMatch searchMatch(
      CatalogLookupRef ref, String normalizedQuery, List<String> tokens) {
    String lookupId = ref.lookupId().toLowerCase(Locale.ROOT);
    String qualifiedId = ref.qualifiedId().toLowerCase(Locale.ROOT);
    String catalogGroup = ref.catalogGroup().toLowerCase(Locale.ROOT);
    String summary = summaryFor(ref.value()).toLowerCase(Locale.ROOT);
    String combined = lookupId + " " + qualifiedId + " " + catalogGroup + " " + summary;
    Integer rank = null;
    if (lookupId.equals(normalizedQuery) || qualifiedId.equals(normalizedQuery)) {
      rank = 0;
    } else if (containsAllTokens(lookupId, tokens) || containsAllTokens(qualifiedId, tokens)) {
      rank = 1;
    } else if (containsAllTokens(summary, tokens)) {
      rank = 2;
    } else if (containsAllTokens(combined, tokens)) {
      rank = 3;
    }
    if (rank == null) {
      return null;
    }
    return new RankedSearchMatch(
        rank,
        new CatalogSearchMatch(
            ref.catalogGroup(),
            ref.lookupId(),
            ref.qualifiedId(),
            kindFor(ref.value()),
            summaryFor(ref.value())));
  }

  private static boolean containsAllTokens(String haystack, List<String> tokens) {
    for (String token : tokens) {
      if (!haystack.contains(token)) {
        return false;
      }
    }
    return true;
  }

  static String kindFor(Object value) {
    return switch (value) {
      case TypeEntry _ -> "ENTRY";
      case NestedTypeGroup _ -> "NESTED_GROUP";
      case PlainTypeGroup _ -> "PLAIN_GROUP";
      default -> throw new IllegalArgumentException("Unsupported catalog lookup value: " + value);
    };
  }

  static String summaryFor(Object value) {
    return switch (value) {
      case TypeEntry entry -> entry.summary();
      case NestedTypeGroup group ->
          "Nested type group with discriminator "
              + group.discriminatorField()
              + " and "
              + group.types().size()
              + " variants.";
      case PlainTypeGroup group -> group.type().summary();
      default -> throw new IllegalArgumentException("Unsupported catalog lookup value: " + value);
    };
  }

  private static List<CatalogEntryRef> matchingEntryRefs(Catalog catalog, String idOrQualifiedId) {
    String lookup =
        CatalogRecordValidation.requireNonBlank(idOrQualifiedId, "idOrQualifiedId").trim();
    int separator = lookup.indexOf(':');
    if (separator >= 0) {
      String group = lookup.substring(0, separator).trim();
      String id = lookup.substring(separator + 1).trim();
      if (group.isEmpty() || id.isEmpty()) {
        return List.of();
      }
      return allEntryRefs(catalog).stream()
          .filter(entryRef -> entryRef.group().equals(group) && entryRef.entry().id().equals(id))
          .toList();
    }
    return allEntryRefs(catalog).stream()
        .filter(entryRef -> entryRef.entry().id().equals(lookup))
        .toList();
  }

  private static List<CatalogLookupRef> matchingLookupRefs(
      Catalog catalog, String idOrQualifiedId) {
    String lookup =
        CatalogRecordValidation.requireNonBlank(idOrQualifiedId, "idOrQualifiedId").trim();
    int separator = lookup.indexOf(':');
    if (separator >= 0) {
      String group = lookup.substring(0, separator).trim();
      String id = lookup.substring(separator + 1).trim();
      if (group.isEmpty() || id.isEmpty()) {
        return List.of();
      }
      return allLookupRefs(catalog).stream()
          .filter(
              lookupRef ->
                  lookupRef.catalogGroup().equals(group) && lookupRef.lookupId().equals(id))
          .toList();
    }
    return allLookupRefs(catalog).stream()
        .filter(lookupRef -> lookupRef.lookupId().equals(lookup))
        .toList();
  }

  private static List<CatalogEntryRef> allEntryRefs(Catalog catalog) {
    return Stream.of(
            Stream.of(new CatalogEntryRef("requestType", catalog.requestType())),
            entryRefs("sourceTypes", catalog.sourceTypes()).stream(),
            entryRefs("persistenceTypes", catalog.persistenceTypes()).stream(),
            entryRefs("stepTypes", catalog.stepTypes()).stream(),
            entryRefs("mutationActionTypes", catalog.mutationActionTypes()).stream(),
            entryRefs("assertionTypes", catalog.assertionTypes()).stream(),
            entryRefs("inspectionQueryTypes", catalog.inspectionQueryTypes()).stream(),
            catalog.nestedTypes().stream()
                .flatMap(group -> entryRefs(group.group(), group.types()).stream()),
            catalog.plainTypes().stream()
                .map(group -> new CatalogEntryRef(group.group(), group.type())))
        .flatMap(Function.identity())
        .toList();
  }

  private static List<CatalogLookupRef> allLookupRefs(Catalog catalog) {
    return Stream.of(
            allEntryRefs(catalog).stream()
                .map(
                    entryRef ->
                        new CatalogLookupRef(
                            entryRef.group(),
                            entryRef.entry().id(),
                            entryRef.qualifiedId(),
                            entryRef.entry())),
            catalog.nestedTypes().stream()
                .map(
                    group ->
                        new CatalogLookupRef(
                            "nestedTypes", group.group(), "nestedTypes:" + group.group(), group)),
            catalog.plainTypes().stream()
                .map(
                    group ->
                        new CatalogLookupRef(
                            "plainTypes", group.group(), "plainTypes:" + group.group(), group)))
        .flatMap(Function.identity())
        .toList();
  }

  private static List<CatalogEntryRef> entryRefs(String group, List<TypeEntry> entries) {
    return entries.stream().map(entry -> new CatalogEntryRef(group, entry)).toList();
  }

  private record CatalogEntryRef(String group, TypeEntry entry) {
    private CatalogEntryRef {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      Objects.requireNonNull(entry, "entry must not be null");
    }

    private String qualifiedId() {
      return group + ":" + entry.id();
    }
  }

  private record CatalogLookupRef(
      String catalogGroup, String lookupId, String qualifiedId, Object value) {
    private CatalogLookupRef {
      catalogGroup = CatalogRecordValidation.requireNonBlank(catalogGroup, "catalogGroup");
      lookupId = CatalogRecordValidation.requireNonBlank(lookupId, "lookupId");
      qualifiedId = CatalogRecordValidation.requireNonBlank(qualifiedId, "qualifiedId");
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  private record RankedSearchMatch(int rank, CatalogSearchMatch match) {
    private RankedSearchMatch {
      Objects.requireNonNull(match, "match must not be null");
    }
  }
}
