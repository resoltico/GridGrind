package dev.erst.gridgrind.contract.catalog;

import java.util.Objects;

/** Lookup value carrying one published nested type group. */
record CatalogNestedGroupLookupValue(NestedTypeGroup group) implements CatalogLookupValue {
  CatalogNestedGroupLookupValue {
    Objects.requireNonNull(group, "group must not be null");
  }
}
