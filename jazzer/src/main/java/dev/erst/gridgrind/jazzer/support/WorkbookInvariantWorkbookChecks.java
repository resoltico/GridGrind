package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.source.*;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import java.nio.file.*;
import java.util.*;

/** Validates structural engine invariants against a live workbook instance. */
final class WorkbookInvariantWorkbookChecks {
  private WorkbookInvariantWorkbookChecks() {}

  static void requireWorkbookShape(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    var workbookSummary =
        ((dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult)
                readExecutor
                    .read(
                        workbook,
                        new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook-shape")))
                    .getFirst())
            .workbook();

    requireEngineWorkbookSummaryShape(workbookSummary);
    workbook
        .sheets()
        .sheetNames()
        .forEach(
            sheetName -> {
              requireEngineSheetSummaryShape(
                  ((dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult)
                          readExecutor
                              .read(
                                  workbook,
                                  new WorkbookReadCommand.GetSheetSummary(
                                      new StepId("sheet-shape-" + sheetName), sheetName))
                              .getFirst())
                      .sheet());
              ((dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.DrawingObjectsResult)
                      readExecutor
                          .read(
                              workbook,
                              new WorkbookReadCommand.GetDrawingObjects(
                                  new StepId("drawing-shape-" + sheetName), sheetName))
                          .getFirst())
                  .drawingObjects()
                  .forEach(WorkbookInvariantWorkbookChecks::requireEngineDrawingObjectShape);
              ((dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult)
                      readExecutor
                          .read(
                              workbook,
                              new WorkbookReadCommand.GetCharts(
                                  new StepId("chart-shape-" + sheetName),
                                  new dev.erst.gridgrind.excel.ExcelChartSelection.AllOnSheet(
                                      sheetName)))
                          .getFirst())
                  .charts()
                  .forEach(WorkbookInvariantWorkbookChecks::requireEngineChartShape);
            });
    ((dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult)
            readExecutor
                .read(
                    workbook,
                    new WorkbookReadCommand.GetPivotTables(
                        new StepId("pivot-shape"),
                        new dev.erst.gridgrind.core.plan.PivotTableSelector.All()))
                .getFirst())
        .pivotTables()
        .forEach(WorkbookInvariantWorkbookChecks::requireEnginePivotTableShape);
  }

  private static void requireEngineWorkbookSummaryShape(
      dev.erst.gridgrind.core.result.WorkbookSummary workbook) {
    WorkbookInvariantEngineShapeChecks.requireEngineWorkbookSummaryShape(workbook);
  }

  private static void requireEngineSheetSummaryShape(
      dev.erst.gridgrind.core.result.SheetSummaryReport sheet) {
    WorkbookInvariantEngineShapeChecks.requireEngineSheetSummaryShape(sheet);
  }

  private static void requireEngineDrawingObjectShape(
      dev.erst.gridgrind.core.model.drawing.DrawingObjectReport drawingObject) {
    WorkbookInvariantEngineShapeChecks.requireEngineDrawingObjectShape(drawingObject);
  }

  private static void requireEngineChartShape(
      dev.erst.gridgrind.core.model.chart.ChartReport chart) {
    WorkbookInvariantEngineShapeChecks.requireEngineChartShape(chart);
  }

  private static void requireEnginePivotTableShape(
      dev.erst.gridgrind.core.model.pivot.PivotTableReport pivotTable) {
    WorkbookInvariantEngineShapeChecks.requireEnginePivotTableShape(pivotTable);
  }
}
