package dev.erst.gridgrind.core.catalog;

/** Closed catalog lookup result, reusing the published entry and group records. */
public sealed interface CatalogLookupValue
    permits TypeEntry, NestedTypeGroup, PlainTypeGroup, TopLevelTypeGroup {}
