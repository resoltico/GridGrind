package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.*;
import dev.erst.gridgrind.contract.step.*;
import dev.erst.gridgrind.excel.*;
import java.nio.file.*;
import java.util.*;

/** Validates structural engine invariants against a live workbook instance. */
final class WorkbookInvariantWorkbookChecks {
  private WorkbookInvariantWorkbookChecks() {}

  static void requireWorkbookShape(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    var workbookSummary =
        ((dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummaryResult)
                readExecutor
                    .read(workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-shape"))
                    .getFirst())
            .workbook();

    requireEngineWorkbookSummaryShape(workbookSummary);
    workbook
        .sheets()
        .sheetNames()
        .forEach(
            sheetName -> {
              requireEngineSheetSummaryShape(
                  ((dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummaryResult)
                          readExecutor
                              .read(
                                  workbook,
                                  new WorkbookReadCommand.GetSheetSummary(
                                      "sheet-shape-" + sheetName, sheetName))
                              .getFirst())
                      .sheet());
              ((dev.erst.gridgrind.excel.WorkbookDrawingResult.DrawingObjectsResult)
                      readExecutor
                          .read(
                              workbook,
                              new WorkbookReadCommand.GetDrawingObjects(
                                  "drawing-shape-" + sheetName, sheetName))
                          .getFirst())
                  .drawingObjects()
                  .forEach(WorkbookInvariantWorkbookChecks::requireEngineDrawingObjectShape);
              ((dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult)
                      readExecutor
                          .read(
                              workbook,
                              new WorkbookReadCommand.GetCharts(
                                  "chart-shape-" + sheetName,
                                  new dev.erst.gridgrind.excel.ExcelChartSelection.AllOnSheet(
                                      sheetName)))
                          .getFirst())
                  .charts()
                  .forEach(WorkbookInvariantWorkbookChecks::requireEngineChartShape);
            });
    ((dev.erst.gridgrind.excel.WorkbookDrawingResult.PivotTablesResult)
            readExecutor
                .read(
                    workbook,
                    new WorkbookReadCommand.GetPivotTables(
                        "pivot-shape",
                        new dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection.All()))
                .getFirst())
        .pivotTables()
        .forEach(WorkbookInvariantWorkbookChecks::requireEnginePivotTableShape);
  }

  private static void requireEngineWorkbookSummaryShape(
      dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary workbook) {
    WorkbookInvariantEngineShapeChecks.requireEngineWorkbookSummaryShape(workbook);
  }

  private static void requireEngineSheetSummaryShape(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary sheet) {
    WorkbookInvariantEngineShapeChecks.requireEngineSheetSummaryShape(sheet);
  }

  private static void requireEngineDrawingObjectShape(
      dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot drawingObject) {
    WorkbookInvariantEngineShapeChecks.requireEngineDrawingObjectShape(drawingObject);
  }

  private static void requireEngineChartShape(dev.erst.gridgrind.excel.ExcelChartSnapshot chart) {
    WorkbookInvariantEngineShapeChecks.requireEngineChartShape(chart);
  }

  private static void requireEnginePivotTableShape(
      dev.erst.gridgrind.excel.pivot.ExcelPivotTableSnapshot pivotTable) {
    WorkbookInvariantEngineShapeChecks.requireEnginePivotTableShape(pivotTable);
  }
}
