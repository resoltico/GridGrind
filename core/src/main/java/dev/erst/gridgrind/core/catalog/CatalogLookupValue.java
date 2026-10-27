package dev.erst.gridgrind.core.catalog;

/** Closed internal value family behind one published catalog lookup surface. */
sealed interface CatalogLookupValue
    permits CatalogEntryLookupValue,
        CatalogNestedGroupLookupValue,
        CatalogPlainGroupLookupValue,
        CatalogTopLevelGroupLookupValue {}
