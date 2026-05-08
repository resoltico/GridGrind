package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.List;

/** Shared descriptor helpers for the nested-type catalog registries. */
final class GridGrindProtocolCatalogNestedTypeGroupSupport {
  private GridGrindProtocolCatalogNestedTypeGroupSupport() {}

  static CatalogNestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
    return CatalogTypeEntryFactory.nestedTypeGroup(group, sealedType, typeDescriptors);
  }

  static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return CatalogTypeEntryFactory.descriptor(recordType, id, summary, optionalFields);
  }

  static <T extends Record & Selector> CatalogTypeDescriptor selectorDescriptor(
      Class<T> recordType, String summary, String... optionalFields) {
    return descriptor(
        recordType, SelectorJsonSupport.typeIdFor(recordType), summary, optionalFields);
  }
}
