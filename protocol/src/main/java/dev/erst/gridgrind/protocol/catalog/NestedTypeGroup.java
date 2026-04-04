package dev.erst.gridgrind.protocol.catalog;

import java.util.List;

/** JSON-serializable named group of nested tagged union variants. */
public record NestedTypeGroup(String group, List<TypeEntry> types) {
  public NestedTypeGroup {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    types = CatalogRecordValidation.copyEntries(types, "types");
  }
}
