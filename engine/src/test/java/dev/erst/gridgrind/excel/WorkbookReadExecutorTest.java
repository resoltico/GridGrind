package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookReadExecutor command dispatch and null handling. */
class WorkbookReadExecutorTest {
  @Test
  void appliesVarargsAndIterableCommandsInOrder() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");
      sheet.setCell("A1", ExcelCellValue.text("Report"));
      sheet.setCell("B2", ExcelCellValue.number(61.0));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B2")));

      WorkbookReadExecutor executor = new WorkbookReadExecutor();

      List<WorkbookReadResult> varargsResults =
          executor.apply(
              workbook,
              new WorkbookReadCommand.GetWorkbookSummary("workbook"),
              new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1")),
              new WorkbookReadCommand.AnalyzeNamedRangeSurface(
                  "ranges", new ExcelNamedRangeSelection.All()));
      List<WorkbookReadResult> iterableResults =
          executor.apply(
              workbook,
              List.of(
                  new WorkbookReadCommand.GetSheetSummary("sheet", "Budget"),
                  new WorkbookReadCommand.GetMergedRegions("merged", "Budget")));

      assertEquals(List.of("workbook", "cells", "ranges"), requestIds(varargsResults));
      assertInstanceOf(WorkbookReadResult.WorkbookSummaryResult.class, varargsResults.get(0));
      assertInstanceOf(WorkbookReadResult.CellsResult.class, varargsResults.get(1));
      assertInstanceOf(WorkbookReadResult.NamedRangeSurfaceResult.class, varargsResults.get(2));
      assertEquals(List.of("sheet", "merged"), requestIds(iterableResults));
      assertInstanceOf(WorkbookReadResult.SheetSummaryResult.class, iterableResults.get(0));
      assertInstanceOf(WorkbookReadResult.MergedRegionsResult.class, iterableResults.get(1));
    }
  }

  @Test
  void rejectsNullWorkbookCommandsAndCommandEntries() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookReadExecutor executor = new WorkbookReadExecutor();

      assertThrows(
          NullPointerException.class, () -> executor.apply(workbook, (WorkbookReadCommand[]) null));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(null, new WorkbookReadCommand.GetWorkbookSummary("workbook")));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(workbook, (Iterable<WorkbookReadCommand>) null));
      assertThrows(
          NullPointerException.class,
          () ->
              executor.apply(
                  workbook,
                  Arrays.asList(new WorkbookReadCommand.GetWorkbookSummary("workbook"), null)));
    }
  }

  private static List<String> requestIds(List<WorkbookReadResult> results) {
    return results.stream().map(WorkbookReadResult::requestId).toList();
  }
}
