package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** Facade over the split nested/plain field-shape descriptor registries. */
final class GridGrindProtocolCatalogFieldGroupSupport {
  private GridGrindProtocolCatalogFieldGroupSupport() {}

  static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
      GridGrindProtocolCatalogNestedTypeGroups.NESTED_TYPE_GROUPS;

  static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      GridGrindProtocolCatalogPlainTypeDescriptors.PLAIN_TYPE_DESCRIPTORS;
}
