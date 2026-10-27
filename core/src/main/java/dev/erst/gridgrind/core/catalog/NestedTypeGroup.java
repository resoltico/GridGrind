package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** JSON-serializable named group of nested tagged union variants. */
public record NestedTypeGroup(String group, String discriminatorField, List<TypeEntry> types) {
  public NestedTypeGroup {
    group = TextValidation.requireNonBlank(group, "group");
    discriminatorField = TextValidation.requireNonBlank(discriminatorField, "discriminatorField");
    types = CatalogRecordValidation.copyEntries(types, "types");
  }
}
