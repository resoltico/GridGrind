package dev.erst.gridgrind.contract.assertion;

import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.NamedRangeReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.SheetSummaryReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
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

  @ProtocolTypeMetadata(
      id = "EXPECT_WORKBOOK_PROTECTION",
      summary = "Require the workbook protection report to match exactly.",
      targetSelectors = {WorkbookSelector.class})
  record WorkbookProtectionFacts(WorkbookProtectionReport protection)
      implements WorkbookFactAssertion {
    public WorkbookProtectionFacts {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_SHEET_STRUCTURE",
      summary = "Require the selected sheet summary report to match exactly.",
      targetSelectors = {SheetSelector.ByName.class})
  record SheetStructureFacts(SheetSummaryReport sheet) implements WorkbookFactAssertion {
    public SheetStructureFacts {
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_NAMED_RANGE_FACTS",
      summary = "Require the selected named-range reports to match exactly and in order.",
      targetSelectors = {NamedRangeSelector.class})
  record NamedRangeFacts(List<NamedRangeReport> namedRanges) implements WorkbookFactAssertion {
    public NamedRangeFacts {
      namedRanges = AssertionSupport.copyNamedRanges(namedRanges, "namedRanges");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_TABLE_FACTS",
      summary = "Require the selected table reports to match exactly and in order.",
      targetSelectors = {TableSelector.class})
  record TableFacts(List<TableEntryReport> tables) implements WorkbookFactAssertion {
    public TableFacts {
      tables = AssertionSupport.copyTables(tables, "tables");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_PIVOT_TABLE_FACTS",
      summary = "Require the selected pivot-table reports to match exactly and in order.",
      targetSelectors = {PivotTableSelector.class})
  record PivotTableFacts(List<PivotTableReport> pivotTables) implements WorkbookFactAssertion {
    public PivotTableFacts {
      pivotTables = AssertionSupport.copyPivotTables(pivotTables, "pivotTables");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_CHART_FACTS",
      summary = "Require the selected chart reports to match exactly and in order.",
      targetSelectors = {ChartSelector.class})
  record ChartFacts(List<ChartReport> charts) implements WorkbookFactAssertion {
    public ChartFacts {
      charts = AssertionSupport.copyCharts(charts, "charts");
    }
  }
}
