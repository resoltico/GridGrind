package dev.erst.gridgrind.contract.catalog;

import java.util.Objects;

/** Lookup value carrying one published type entry. */
record CatalogEntryLookupValue(TypeEntry entry) implements CatalogLookupValue {
  CatalogEntryLookupValue {
    Objects.requireNonNull(entry, "entry must not be null");
  }
}
