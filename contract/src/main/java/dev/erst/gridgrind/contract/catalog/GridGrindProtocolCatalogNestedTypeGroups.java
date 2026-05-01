package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.stream.Stream;

/** Owns the nested field-shape descriptor registry used by the protocol catalog. */
final class GridGrindProtocolCatalogNestedTypeGroups {
  private GridGrindProtocolCatalogNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
      Stream.of(
              GridGrindProtocolCatalogSourceAndReportNestedTypeGroups.SOURCE_AND_REPORT_GROUPS,
              GridGrindProtocolCatalogWorkbookInputNestedTypeGroups.WORKBOOK_INPUT_GROUPS,
              GridGrindProtocolCatalogSelectorNestedTypeGroups.SELECTOR_GROUPS,
              GridGrindProtocolCatalogChartAndValidationNestedTypeGroups
                  .CHART_AND_VALIDATION_GROUPS)
          .flatMap(List::stream)
          .toList();
}
