package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import java.util.List;
import java.util.Objects;

/** Assertions over workbook, sheet, named-range, table, pivot, and chart fact reports. */
public sealed interface WorkbookFactAssertion extends Assertion
    permits WorkbookFactAssertion.WorkbookProtectionFacts,
        WorkbookFactAssertion.SheetStructureFacts,
        WorkbookFactAssertion.NamedRangeFacts,
        WorkbookFactAssertion.TableFacts,
        WorkbookFactAssertion.PivotTableFacts,
        WorkbookFactAssertion.ChartFacts {

  @Variant(
      id = "EXPECT_WORKBOOK_PROTECTION",
      summary = "Require the workbook protection report to match exactly.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record WorkbookProtectionFacts(WorkbookProtectionReport protection)
      implements WorkbookFactAssertion {
    public WorkbookProtectionFacts {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  @Variant(
      id = "EXPECT_SHEET_STRUCTURE",
      summary = "Require the selected sheet summary report to match exactly.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record SheetStructureFacts(SheetSummaryReport sheet) implements WorkbookFactAssertion {
    public SheetStructureFacts {
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  @Variant(
      id = "EXPECT_NAMED_RANGE_FACTS",
      summary = "Require the selected named-range reports to match exactly and in order.",
      targetSelectors = {NamedRangeSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record NamedRangeFacts(List<NamedRangeReport> namedRanges) implements WorkbookFactAssertion {
    public NamedRangeFacts {
      namedRanges = AssertionSupport.copyNamedRanges(namedRanges, "namedRanges");
    }
  }

  @Variant(
      id = "EXPECT_TABLE_FACTS",
      summary = "Require the selected table reports to match exactly and in order.",
      targetSelectors = {TableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record TableFacts(List<TableEntryReport> tables) implements WorkbookFactAssertion {
    public TableFacts {
      tables = AssertionSupport.copyTables(tables, "tables");
    }
  }

  @Variant(
      id = "EXPECT_PIVOT_TABLE_FACTS",
      summary = "Require the selected pivot-table reports to match exactly and in order.",
      targetSelectors = {PivotTableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record PivotTableFacts(List<PivotTableReport> pivotTables) implements WorkbookFactAssertion {
    public PivotTableFacts {
      pivotTables = AssertionSupport.copyPivotTables(pivotTables, "pivotTables");
    }
  }

  @Variant(
      id = "EXPECT_CHART_FACTS",
      summary = "Require the selected chart reports to match exactly and in order.",
      targetSelectors = {ChartSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record ChartFacts(List<ChartReport> charts) implements WorkbookFactAssertion {
    public ChartFacts {
      charts = AssertionSupport.copyCharts(charts, "charts");
    }
  }
}
