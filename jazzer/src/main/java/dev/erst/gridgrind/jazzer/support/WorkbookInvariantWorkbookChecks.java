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
                        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                            new StepId("workbook-shape"),
                            new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                            new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                                .GetWorkbookSummary()))
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
                                  new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                      new StepId("sheet-shape-" + sheetName),
                                      new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                                          new dev.erst.gridgrind.core.vocabulary.SheetName(
                                              sheetName)),
                                      new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                                          .GetSheetSummary()))
                              .getFirst())
                      .sheet());
              ((dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.DrawingObjectsResult)
                      readExecutor
                          .read(
                              workbook,
                              new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                  new StepId("drawing-shape-" + sheetName),
                                  new dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet(
                                      new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName)),
                                  new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                                      .GetDrawingObjects()))
                          .getFirst())
                  .drawingObjects()
                  .forEach(WorkbookInvariantWorkbookChecks::requireEngineDrawingObjectShape);
              ((dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult)
                      readExecutor
                          .read(
                              workbook,
                              new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                  new StepId("chart-shape-" + sheetName),
                                  new dev.erst.gridgrind.core.plan.ChartSelector.AllOnSheet(
                                      new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName)),
                                  new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                                      .GetCharts()))
                          .getFirst())
                  .charts()
                  .forEach(WorkbookInvariantWorkbookChecks::requireEngineChartShape);
            });
    ((dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult)
            readExecutor
                .read(
                    workbook,
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("pivot-shape"),
                        new dev.erst.gridgrind.core.plan.PivotTableSelector.All(),
                        new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                            .GetPivotTables()))
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
