package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.WorkbookSheetResult;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/** Owns expected sheet-layout state captured before `.xlsx` round-tripping. */
final class XlsxRoundTripExpectedLayoutSupport {
  private XlsxRoundTripExpectedLayoutSupport() {}

  static Map<String, XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState>
      expectedSheetLayouts(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState>
        expectedLayouts = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      WorkbookSheetResult.SheetLayoutResult layoutResult =
          (WorkbookSheetResult.SheetLayoutResult)
              readExecutor
                  .apply(workbook, new WorkbookReadCommand.GetSheetLayout("layout", sheetName))
                  .getFirst();
      expectedLayouts.put(sheetName, expectedSheetLayout(layoutResult.layout()));
    }
    return Map.copyOf(expectedLayouts);
  }

  static XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState expectedSheetLayout(
      WorkbookSheetResult.SheetLayout layout) {
    Objects.requireNonNull(layout, "layout must not be null");

    LinkedHashMap<Integer, XlsxRoundTripExpectedStateSupport.ExpectedRowLayoutState> expectedRows =
        new LinkedHashMap<>();
    for (WorkbookSheetResult.RowLayout row : layout.rows()) {
      expectedRows.put(
          row.rowIndex(),
          new XlsxRoundTripExpectedStateSupport.ExpectedRowLayoutState(
              row.hidden(), row.outlineLevel(), row.collapsed()));
    }

    LinkedHashMap<Integer, XlsxRoundTripExpectedStateSupport.ExpectedColumnLayoutState>
        expectedColumns = new LinkedHashMap<>();
    for (WorkbookSheetResult.ColumnLayout column : layout.columns()) {
      expectedColumns.put(
          column.columnIndex(),
          new XlsxRoundTripExpectedStateSupport.ExpectedColumnLayoutState(
              column.hidden(), column.outlineLevel(), column.collapsed()));
    }

    return new XlsxRoundTripExpectedStateSupport.ExpectedSheetLayoutState(
        layout.pane(), layout.zoomPercent(), Map.copyOf(expectedRows), Map.copyOf(expectedColumns));
  }
}
