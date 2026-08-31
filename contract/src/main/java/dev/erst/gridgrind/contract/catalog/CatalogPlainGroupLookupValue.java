package dev.erst.gridgrind.contract.catalog;

import java.util.Objects;

/** Lookup value carrying one published plain type group. */
record CatalogPlainGroupLookupValue(PlainTypeGroup group) implements CatalogLookupValue {
  CatalogPlainGroupLookupValue {
    Objects.requireNonNull(group, "group must not be null");
  }
}
