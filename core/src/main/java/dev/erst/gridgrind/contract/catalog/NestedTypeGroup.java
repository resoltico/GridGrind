package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** JSON-serializable named group of nested tagged union variants. */
public record NestedTypeGroup(String group, String discriminatorField, List<TypeEntry> types) {
  public NestedTypeGroup {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    discriminatorField =
        CatalogRecordValidation.requireNonBlank(discriminatorField, "discriminatorField");
    types = CatalogRecordValidation.copyEntries(types, "types");
  }
}
