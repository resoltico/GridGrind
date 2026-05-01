package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.junit.jupiter.api.Test;

/** Pivot-table deletion failure and shared-cache coverage. */
class ExcelPivotTableDeletionCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void deletePivotTableReportsRelationRemovalFailure() throws Exception {
    try (ExcelWorkbook workbook = pivotWorkbook()) {
      ExcelPivotTableController failingController =
          new ExcelPivotTableController((parent, child) -> false);

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> failingController.deletePivotTable(workbook, "Sales Pivot", "Report"));
      assertTrue(failure.getMessage().contains("Failed to remove pivot table relation"));
    }
  }

  @Test
  void deletePivotTableReportsCacheDefinitionRemovalFailure() throws Exception {
    try (ExcelWorkbook workbook = pivotWorkbook()) {
      AtomicInteger calls = new AtomicInteger();
      ExcelPivotTableController failingController =
          new ExcelPivotTableController((parent, child) -> calls.incrementAndGet() == 1);

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> failingController.deletePivotTable(workbook, "Sales Pivot", "Report"));
      assertTrue(failure.getMessage().contains("pivot cache definition relation"));
    }
  }

  @Test
  void sharedCacheAndReadbackEdgeCasesCoverRemainingControllerBranches() throws Exception {
    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable firstPivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotTable secondPivot = pivotWorkbookWithSecondPivot(workbook);
      secondPivot.setPivotCacheDefinition(firstPivot.getPivotCacheDefinition());

      invokeVoid(controller, "deletePivotHandle", workbook, allPivotHandles(workbook).getFirst());
      assertEquals(1, workbook.xssfWorkbook().getPivotTables().size());
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable firstPivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotTable secondPivot = pivotWorkbookWithSecondPivot(workbook);
      secondPivot.setPivotCacheDefinition(null);

      assertFalse(
          invoke(
              controller,
              "cacheDefinitionShared",
              Boolean.class,
              workbook,
              allPivotHandles(workbook).getFirst(),
              firstPivot.getPivotCacheDefinition()));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      controller.setPivotTable(
          workbook,
          definition(
              "Filter Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "C5",
              List.of("Owner"),
              List.of(),
              List.of("Stage")));

      ExcelPivotTableSnapshot.Supported snapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Supported.class,
              controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst());
      assertEquals(List.of(), fieldNames(snapshot.rowLabels()));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "DupReadSource",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      controller.setPivotTable(
          workbook,
          definition(
              "Named Read Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.NamedRange("DupReadSource"),
              "C5",
              List.of(),
              List.of("Region"),
              List.of()));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "DupReadSource",
              new ExcelNamedRangeScope.SheetScope("Data"),
              new ExcelNamedRangeTarget("Data", "A1:D5")));

      IllegalArgumentException failure =
          assertInvocationFailure(
              IllegalArgumentException.class,
              () ->
                  invoke(
                      controller,
                      "snapshotSource",
                      ExcelPivotTableSnapshot.Source.class,
                      workbook.xssfWorkbook(),
                      workbook.xssfWorkbook().getPivotTables().getFirst()));
      assertTrue(failure.getMessage().contains("ambiguous"));
    }
  }
}
