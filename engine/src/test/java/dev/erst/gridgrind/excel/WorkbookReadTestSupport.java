package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;

/** Shared defaults for read-command and read-result tests after projection became explicit. */
final class WorkbookReadTestSupport {
  private WorkbookReadTestSupport() {}

  static CellReadProjection projection() {
    return CellReadProjection.defaults();
  }

  static WorkbookReadCommand.GetCells getCells(
      String stepId, String sheetName, List<String> addresses) {
    return new WorkbookReadCommand.GetCells(new StepId(stepId), sheetName, addresses, projection());
  }

  static WorkbookReadCommand.GetWindow getWindow(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WorkbookReadCommand.GetWindow(
        new StepId(stepId),
        sheetName,
        new ExcelReadWindow(topLeftAddress, rowCount, columnCount),
        projection(),
        false);
  }

  static WorkbookReadCommand.GetSheetSchema getSheetSchema(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WorkbookReadCommand.GetSheetSchema(
        new StepId(stepId),
        sheetName,
        new ExcelReadWindow(topLeftAddress, rowCount, columnCount),
        projection());
  }

  static dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult cellsResult(
      String stepId, String sheetName, List<CellReport> cells) {
    return new dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult(
        new StepId(stepId), new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), cells);
  }

  static dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.SheetSchemaResult
      sheetSchemaResult(String stepId, SheetSchemaReport surface) {
    return new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.SheetSchemaResult(
        new StepId(stepId), surface);
  }
}
