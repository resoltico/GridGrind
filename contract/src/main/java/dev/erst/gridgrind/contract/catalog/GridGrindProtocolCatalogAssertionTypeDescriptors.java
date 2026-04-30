package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.assertion.Assertion;
import java.util.List;

/** Assertion descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogAssertionTypeDescriptors {
  static final List<CatalogTypeDescriptor> ASSERTION_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              Assertion.NamedRangePresent.class,
              "EXPECT_NAMED_RANGE_PRESENT",
              "Require the selected named-range selector to resolve to one or more named"
                  + " ranges."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.NamedRangeAbsent.class,
              "EXPECT_NAMED_RANGE_ABSENT",
              "Require the selected named-range selector to resolve to no named ranges."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.TablePresent.class,
              "EXPECT_TABLE_PRESENT",
              "Require the selected table selector to resolve to one or more tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.TableAbsent.class,
              "EXPECT_TABLE_ABSENT",
              "Require the selected table selector to resolve to no tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.PivotTablePresent.class,
              "EXPECT_PIVOT_TABLE_PRESENT",
              "Require the selected pivot-table selector to resolve to one or more pivot"
                  + " tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.PivotTableAbsent.class,
              "EXPECT_PIVOT_TABLE_ABSENT",
              "Require the selected pivot-table selector to resolve to no pivot tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.ChartPresent.class,
              "EXPECT_CHART_PRESENT",
              "Require the selected chart selector to resolve to one or more charts."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.ChartAbsent.class,
              "EXPECT_CHART_ABSENT",
              "Require the selected chart selector to resolve to no charts."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.CellValue.class,
              "EXPECT_CELL_VALUE",
              "Require every selected cell to have the exact effective value."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.DisplayValue.class,
              "EXPECT_DISPLAY_VALUE",
              "Require every selected cell to have the exact formatted display string."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.FormulaText.class,
              "EXPECT_FORMULA_TEXT",
              "Require every selected cell to store the exact formula text."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.CellStyle.class,
              "EXPECT_CELL_STYLE",
              "Require every selected cell to have the exact style snapshot."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.WorkbookProtectionFacts.class,
              "EXPECT_WORKBOOK_PROTECTION",
              "Require the workbook protection report to match exactly."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.SheetStructureFacts.class,
              "EXPECT_SHEET_STRUCTURE",
              "Require the selected sheet summary report to match exactly."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.NamedRangeFacts.class,
              "EXPECT_NAMED_RANGE_FACTS",
              "Require the selected named-range reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.TableFacts.class,
              "EXPECT_TABLE_FACTS",
              "Require the selected table reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.PivotTableFacts.class,
              "EXPECT_PIVOT_TABLE_FACTS",
              "Require the selected pivot-table reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.ChartFacts.class,
              "EXPECT_CHART_FACTS",
              "Require the selected chart reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnalysisMaxSeverity.class,
              "EXPECT_ANALYSIS_MAX_SEVERITY",
              "Run one analysis query against the selected target and require its highest finding"
                  + " severity to be no higher than maximumSeverity."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnalysisFindingPresent.class,
              "EXPECT_ANALYSIS_FINDING_PRESENT",
              "Run one analysis query against the selected target and require at least one"
                  + " matching finding. severity and messageContains are optional match"
                  + " refinements.",
              "severity",
              "messageContains"),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnalysisFindingAbsent.class,
              "EXPECT_ANALYSIS_FINDING_ABSENT",
              "Run one analysis query against the selected target and require no matching"
                  + " finding. severity and messageContains are optional match refinements.",
              "severity",
              "messageContains"),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AllOf.class,
              "ALL_OF",
              "Require every nested assertion to pass against the same step target."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnyOf.class,
              "ANY_OF",
              "Require at least one nested assertion to pass against the same step target."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.Not.class,
              "NOT",
              "Invert one nested assertion against the same step target."));

  private GridGrindProtocolCatalogAssertionTypeDescriptors() {}
}
