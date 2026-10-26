package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.result.WorkbookSummary;

/** Converts workbook-global security, summary, protection, and named-range snapshots. */
final class InspectionResultWorkbookCoreReportSupport {
  private InspectionResultWorkbookCoreReportSupport() {}

  static WorkbookSummary toWorkbookSummary(
      dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary workbookSummary) {
    return switch (workbookSummary) {
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.Empty empty ->
          new WorkbookSummary.Empty(
              empty.sheetCount(),
              empty.sheetNames().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                  .toList(),
              empty.namedRangeCount(),
              empty.forceFormulaRecalculationOnOpen());
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets withSheets ->
          new WorkbookSummary.WithSheets(
              withSheets.sheetCount(),
              withSheets.sheetNames().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                  .toList(),
              new dev.erst.gridgrind.core.vocabulary.SheetLabel(withSheets.activeSheetName()),
              withSheets.selectedSheetNames().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                  .toList(),
              withSheets.namedRangeCount(),
              withSheets.forceFormulaRecalculationOnOpen());
    };
  }
}
