package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/**
 * One serializable top-level catalog category returned by a {@code --lookup <categoryName>} call.
 */
public record TopLevelTypeGroup(String group, List<TypeEntry> types) {
  public TopLevelTypeGroup {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    types = CatalogRecordValidation.copyEntries(types, "types");
  }
}
