package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** Lookup value carrying one published top-level operation category. */
record CatalogTopLevelGroupLookupValue(String group, List<TypeEntry> types)
    implements CatalogLookupValue {
  CatalogTopLevelGroupLookupValue {
    Objects.requireNonNull(group, "group must not be null");
    types = List.copyOf(Objects.requireNonNull(types, "types must not be null"));
  }
}
