package dev.erst.gridgrind.contract.catalog;

import java.util.ArrayList;
import java.util.List;

/** Facade over the split plain field-shape descriptor registries. */
final class GridGrindProtocolCatalogPlainTypeDescriptors {
  private GridGrindProtocolCatalogPlainTypeDescriptors() {}

  static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS = descriptors();

  private static List<CatalogPlainTypeDescriptor> descriptors() {
    List<CatalogPlainTypeDescriptor> descriptors = new ArrayList<>();
    descriptors.addAll(GridGrindProtocolCatalogExecutionPlainTypeDescriptors.DESCRIPTORS);
    descriptors.addAll(GridGrindProtocolCatalogDrawingPlainTypeDescriptors.DESCRIPTORS);
    descriptors.addAll(GridGrindProtocolCatalogWorkbookAuthoringPlainTypeDescriptors.DESCRIPTORS);
    descriptors.addAll(GridGrindProtocolCatalogWorkbookReportPlainTypeDescriptors.DESCRIPTORS);
    return List.copyOf(descriptors);
  }
}
