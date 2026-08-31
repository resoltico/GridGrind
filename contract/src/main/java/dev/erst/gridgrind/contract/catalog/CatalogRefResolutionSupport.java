package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Stream;

/** Catalog ref building: flat entry and lookup ref lists and id-based matching. */
final class CatalogRefResolutionSupport {
  private CatalogRefResolutionSupport() {}

  static List<CatalogEntryRef> allEntryRefs(Catalog catalog) {
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

  static List<CatalogEntryRef> topLevelEntryRefs(Catalog catalog) {
    return Stream.of(
            entryRefs("mutationActionTypes", catalog.mutationActionTypes()).stream(),
            entryRefs("assertionTypes", catalog.assertionTypes()).stream(),
            entryRefs("inspectionQueryTypes", catalog.inspectionQueryTypes()).stream())
        .flatMap(Function.identity())
        .toList();
  }

  static List<CatalogLookupRef> allLookupRefs(Catalog catalog) {
    return Stream.of(
            allEntryRefs(catalog).stream()
                .map(
                    entryRef ->
                        new CatalogLookupRef(
                            entryRef.group(),
                            entryRef.entry().id(),
                            entryRef.qualifiedId(),
                            new CatalogEntryLookupValue(entryRef.entry()))),
            catalog.nestedTypes().stream()
                .map(
                    group ->
                        new CatalogLookupRef(
                            "nestedTypes",
                            group.group(),
                            "nestedTypes:" + group.group(),
                            new CatalogNestedGroupLookupValue(group))),
            catalog.plainTypes().stream()
                .map(
                    group ->
                        new CatalogLookupRef(
                            "plainTypes",
                            group.group(),
                            "plainTypes:" + group.group(),
                            new CatalogPlainGroupLookupValue(group))),
            topLevelGroupRefs(catalog).stream())
        .flatMap(Function.identity())
        .toList();
  }

  private static List<CatalogLookupRef> topLevelGroupRefs(Catalog catalog) {
    return catalog.topLevelGroups().stream()
        .map(group -> topLevelGroupRef(group.group(), group.types()))
        .toList();
  }

  private static CatalogLookupRef topLevelGroupRef(String group, List<TypeEntry> types) {
    return new CatalogLookupRef(
        group, group, group, new CatalogTopLevelGroupLookupValue(group, types));
  }

  static List<CatalogEntryRef> matchingEntryRefs(Catalog catalog, String idOrQualifiedId) {
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

  static List<CatalogLookupRef> matchingLookupRefs(Catalog catalog, String idOrQualifiedId) {
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

  private static List<CatalogEntryRef> entryRefs(String group, List<TypeEntry> entries) {
    return entries.stream().map(entry -> new CatalogEntryRef(group, entry)).toList();
  }

  /** Internal typed carrier for one catalog entry in one group. */
  record CatalogEntryRef(String group, TypeEntry entry) {
    CatalogEntryRef {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      Objects.requireNonNull(entry, "entry must not be null");
    }

    String qualifiedId() {
      return group + ":" + entry.id();
    }
  }
}
