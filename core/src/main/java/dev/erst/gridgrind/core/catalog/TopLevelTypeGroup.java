package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** One serializable top-level catalog category addressable by its published name. */
public record TopLevelTypeGroup(String group, List<TypeEntry> types) {
  public TopLevelTypeGroup {
    group = TextValidation.requireNonBlank(group, "group");
    types = CatalogRecordValidation.copyEntries(types, "types");
  }
}
