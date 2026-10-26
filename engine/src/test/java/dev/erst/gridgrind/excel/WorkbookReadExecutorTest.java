package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookReadExecutor command dispatch and null handling. */
class WorkbookReadExecutorTest {
  @Test
  void windowQueriesSelectDensityAndPreserveDenseBlankMetadata() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Grid");
      sheet.cells().setCell("A1", new dev.erst.gridgrind.core.model.value.CellValue.Text("header"));
      sheet.cells().setCell("B2", new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(3));
      sheet.cells().setCell("B3", new dev.erst.gridgrind.core.model.formula.ResolvedFormula("1+1"));
      sheet
          .annotations()
          .setComment(
              "B1",
              new dev.erst.gridgrind.core.model.drawing.ResolvedComment(
                  "blank metadata", "tester", false));
      var projection =
          new dev.erst.gridgrind.core.plan.CellReadProjection(
              List.of(dev.erst.gridgrind.core.plan.CellReadFacet.values()));
      var results =
          new WorkbookExecutionEngine()
              .read(
                  workbook,
                  new WorkbookReadCommand.GetWindow(
                      new StepId("sparse"),
                      "Grid",
                      new ExcelReadWindow("A1", 3, 2),
                      projection,
                      false),
                  new WorkbookReadCommand.GetWindow(
                      new StepId("dense"),
                      "Grid",
                      new ExcelReadWindow("A1", 3, 2),
                      projection,
                      true));
      var sparseResult =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.WindowResult.class,
              results.get(0));
      var denseResult =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.WindowResult.class,
              results.get(1));
      assertEquals("sparse", sparseResult.stepId().value());
      assertEquals("dense", denseResult.stepId().value());
      var sparse =
          assertInstanceOf(
              dev.erst.gridgrind.core.model.layout.WindowReport.Sparse.class,
              sparseResult.window());
      var dense =
          assertInstanceOf(
              dev.erst.gridgrind.core.model.layout.WindowReport.Dense.class, denseResult.window());
      assertEquals(
          List.of("A1", "B2", "B3"),
          sparse.populatedCells().stream().map(cell -> cell.address().value()).toList());
      assertEquals(3, dense.dimensions().rowCount());
      assertEquals(2, dense.dimensions().columnCount());
      assertEquals(3, dense.rows().size());
      assertTrue(dense.rows().stream().allMatch(row -> row.cells().size() == 2));
      var blank =
          assertInstanceOf(
              dev.erst.gridgrind.core.model.value.CellReport.BlankReport.class,
              dense.rows().getFirst().cells().get(1));
      assertEquals("B1", blank.address().value());
      assertEquals("blank metadata", blank.comment().orElseThrow().text());
    }
  }

  @Test
  void appliesVarargsAndIterableCommandsInOrder() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");
      sheet.cells().setCell("A1", new CellValue.Text("Report"));
      sheet.cells().setCell("B2", new CellValue.NumberValue(61.0));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B2"))));

      WorkbookReadExecutor executor = new WorkbookReadExecutor();

      List<InspectionResult> varargsResults =
          executor.apply(
              workbook,
              new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook")),
              WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1")),
              new WorkbookReadCommand.GetNamedRangeSurface(
                  new StepId("ranges"), new ExcelNamedRangeSelection.All()));
      List<InspectionResult> iterableResults =
          executor.apply(
              workbook,
              List.of(
                  new WorkbookReadCommand.GetSheetSummary(new StepId("sheet"), "Budget"),
                  new WorkbookReadCommand.GetMergedRegions(new StepId("merged"), "Budget")));
      List<InspectionResult> explicitLocationResults =
          executor.apply(
              workbook,
              new WorkbookLocation.UnsavedWorkbook(),
              WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 1, 1));

      assertEquals(List.of("workbook", "cells", "ranges"), stepIds(varargsResults));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult.class,
          varargsResults.get(0));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult.class,
          varargsResults.get(1));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult
              .class,
          varargsResults.get(2));
      assertEquals(List.of("sheet", "merged"), stepIds(iterableResults));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult.class,
          iterableResults.get(0));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.SheetInspectionResult.MergedRegionsResult.class,
          iterableResults.get(1));
      assertEquals(List.of("window"), stepIds(explicitLocationResults));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.SheetInspectionResult.WindowResult.class,
          explicitLocationResults.getFirst());
    }
  }

  @Test
  void rejectsNullWorkbookCommandsAndCommandEntries() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookReadExecutor executor = new WorkbookReadExecutor();

      assertThrows(
          NullPointerException.class, () -> executor.apply(workbook, (WorkbookReadCommand[]) null));
      assertThrows(
          NullPointerException.class,
          () ->
              executor.apply(
                  null, new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook"))));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(workbook, (Iterable<WorkbookReadCommand>) null));
      assertThrows(
          NullPointerException.class,
          () ->
              executor.apply(
                  workbook,
                  Arrays.asList(
                      new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook")), null)));
      assertThrows(
          NullPointerException.class,
          () ->
              executor.apply(
                  workbook,
                  (WorkbookLocation) null,
                  new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook"))));

      List<InspectionResult> analysisResults =
          executor.apply(
              workbook,
              new WorkbookReadCommand.AnalyzeFormulaHealth(
                  new StepId("formulaHealth"), new ExcelSheetSelection.All()),
              new WorkbookReadCommand.AnalyzeWorkbookFindings(new StepId("workbookFindings")));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult.class,
          analysisResults.get(0));
      assertInstanceOf(
          dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult.class,
          analysisResults.get(1));
    }
  }

  private static List<String> stepIds(List<InspectionResult> results) {
    return results.stream().map(result -> result.stepId().value()).toList();
  }
}
