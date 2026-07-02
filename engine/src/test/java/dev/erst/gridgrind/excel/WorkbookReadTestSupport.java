package dev.erst.gridgrind.excel;

import java.util.List;

/** Shared defaults for read-command and read-result tests after projection became explicit. */
final class WorkbookReadTestSupport {
  private WorkbookReadTestSupport() {}

  static ExcelCellReadProjection projection() {
    return ExcelCellReadProjection.defaults();
  }

  static WorkbookReadCommand.GetCells getCells(
      String stepId, String sheetName, List<String> addresses) {
    return new WorkbookReadCommand.GetCells(stepId, sheetName, addresses, projection());
  }

  static WorkbookReadCommand.GetWindow getWindow(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WorkbookReadCommand.GetWindow(
        stepId,
        sheetName,
        new ExcelReadWindow(topLeftAddress, rowCount, columnCount),
        projection(),
        false);
  }

  static WorkbookReadCommand.GetSheetSchema getSheetSchema(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WorkbookReadCommand.GetSheetSchema(
        stepId,
        sheetName,
        new ExcelReadWindow(topLeftAddress, rowCount, columnCount),
        projection());
  }

  static WorkbookSheetResult.CellsResult cellsResult(
      String stepId, String sheetName, List<ExcelCellSnapshot> cells) {
    return new WorkbookSheetResult.CellsResult(stepId, sheetName, cells, projection(), false);
  }

  static WorkbookSheetResult.WindowResult windowResult(
      String stepId, WorkbookSheetResult.Window window) {
    return new WorkbookSheetResult.WindowResult(stepId, window, projection(), false, false);
  }

  static WorkbookSurfaceResult.SheetSchemaResult sheetSchemaResult(
      String stepId, WorkbookSurfaceResult.SheetSchema surface) {
    return new WorkbookSurfaceResult.SheetSchemaResult(stepId, surface, projection(), false);
  }
}
