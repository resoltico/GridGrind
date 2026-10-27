package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/** Owns expected sheet-layout state captured before `.xlsx` round-tripping. */
final class XlsxRoundTripExpectedLayoutSupport {
  private XlsxRoundTripExpectedLayoutSupport() {}

  static Map<String, XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState>
      expectedSheetLayouts(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    LinkedHashMap<String, XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState>
        expectedLayouts = new LinkedHashMap<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetLayoutResult layoutResult =
          (dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetLayoutResult)
              readExecutor
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("layout"),
                          new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                              new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName)),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetSheetLayout()))
                  .getFirst();
      expectedLayouts.put(sheetName, expectedSheetLayout(layoutResult.layout()));
    }
    return Map.copyOf(expectedLayouts);
  }

  static XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState expectedSheetLayout(
      SheetLayoutReport layout) {
    Objects.requireNonNull(layout, "layout must not be null");

    LinkedHashMap<Integer, XlsxRoundTripExpectedStateSupport.ExpectedRowLayoutState> expectedRows =
        new LinkedHashMap<>();
    for (RowLayoutReport row : layout.rows()) {
      expectedRows.put(
          row.rowIndex(),
          new XlsxRoundTripExpectedStateSupport.ExpectedRowLayoutState(
              row.hidden(), row.outlineLevel(), row.collapsed()));
    }

    LinkedHashMap<Integer, XlsxRoundTripExpectedStateSupport.ExpectedColumnLayoutState>
        expectedColumns = new LinkedHashMap<>();
    for (ColumnLayoutReport column : layout.columns()) {
      expectedColumns.put(
          column.columnIndex(),
          new XlsxRoundTripExpectedStateSupport.ExpectedColumnLayoutState(
              column.hidden(), column.outlineLevel(), column.collapsed()));
    }

    return new XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState(
        layout.pane(), layout.zoomPercent(), Map.copyOf(expectedRows), Map.copyOf(expectedColumns));
  }
}
