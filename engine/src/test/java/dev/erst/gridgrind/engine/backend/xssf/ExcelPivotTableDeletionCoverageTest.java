package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.engine.pivots.PivotAdapter;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.junit.jupiter.api.Test;

/** Pivot-table deletion failure and shared-cache coverage. */
class ExcelPivotTableDeletionCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void deletePivotTableReportsRelationRemovalFailure() throws Exception {
    try (XssfWorkbook workbook = pivotWorkbook()) {
      PivotAdapter failingController = new PivotAdapter((parent, child) -> false);

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> failingController.deletePivotTable(workbook, "Sales Pivot", "Report"));
      assertTrue(failure.getMessage().contains("Failed to remove pivot table relation"));
    }
  }

  @Test
  void deletePivotTableReportsCacheDefinitionRemovalFailure() throws Exception {
    try (XssfWorkbook workbook = pivotWorkbook()) {
      AtomicInteger calls = new AtomicInteger();
      PivotAdapter failingController =
          new PivotAdapter((parent, child) -> calls.incrementAndGet() == 1);

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> failingController.deletePivotTable(workbook, "Sales Pivot", "Report"));
      assertTrue(failure.getMessage().contains("pivot cache definition relation"));
    }
  }

  @Test
  void sharedCacheAndReadbackEdgeCasesCoverRemainingControllerBranches() throws Exception {
    try (XssfWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable firstPivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotTable secondPivot = pivotWorkbookWithSecondPivot(workbook);
      secondPivot.setPivotCacheDefinition(firstPivot.getPivotCacheDefinition());

      invokeVoid(controller, "deletePivotHandle", workbook, allPivotHandles(workbook).getFirst());
      assertEquals(1, workbook.xssfWorkbook().getPivotTables().size());
    }

    try (XssfWorkbook workbook = pivotWorkbook()) {
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

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      controller.setPivotTable(
          workbook,
          definition(
              "Filter Pivot",
              "Report",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "C5",
              List.of("Owner"),
              List.of(),
              List.of("Stage")));

      PivotTableReport.Supported snapshot =
          assertInstanceOf(
              PivotTableReport.Supported.class,
              controller.pivotTables(workbook, new PivotTableSelector.All()).getFirst());
      assertEquals(List.of(), fieldNames(snapshot.rowLabels()));
    }

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("DupReadSource"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5"))));
      controller.setPivotTable(
          workbook,
          definition(
              "Named Read Pivot",
              "Report",
              new PivotTableSpec.PivotSourceSpec.NamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("DupReadSource")),
              "C5",
              List.of(),
              List.of("Region"),
              List.of()));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("DupReadSource"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data")),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5"))));

      IllegalArgumentException failure =
          assertInvocationFailure(
              IllegalArgumentException.class,
              () ->
                  invoke(
                      controller,
                      "snapshotSource",
                      PivotTableReport.PivotSourceReport.class,
                      workbook.xssfWorkbook(),
                      workbook.xssfWorkbook().getPivotTables().getFirst()));
      assertTrue(failure.getMessage().contains("ambiguous"));
    }
  }
}
