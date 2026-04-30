package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.junit.jupiter.api.Test;

/** Pivot-table health and malformed-state coverage. */
class ExcelPivotTableHealthCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void pivotTableHealthCoversMalformedNamesLocationsAndCacheRelations() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      workbook.getOrCreateSheet("OtherReport");
      controller.setPivotTable(
          workbook,
          definition(
              "Alpha Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "C5",
              List.of(),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Beta Pivot",
              "OtherReport",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "C5",
              List.of(),
              List.of("Stage"),
              List.of()));

      XSSFPivotTable firstPivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotTable secondPivot = workbook.xssfWorkbook().getPivotTables().get(1);
      secondPivot.getCTPivotTableDefinition().setName("Alpha Pivot");

      List<AnalysisFindingCode> duplicateCodes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(duplicateCodes.contains(AnalysisFindingCode.PIVOT_TABLE_DUPLICATE_NAME));

      firstPivot.getCTPivotTableDefinition().setName(null);
      ExcelPivotTableSnapshot syntheticNameSnapshot =
          controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst();
      assertTrue(syntheticNameSnapshot.name().startsWith("_GG_PIVOT_"));

      Object firstHandle = allPivotHandles(workbook).getFirst();
      firstPivot.getCTPivotTableDefinition().getLocation().setRef(null);
      assertNull(invoke(controller, "rawLocationRange", Object.class, firstHandle));
      assertEquals(Optional.empty(), invoke(controller, "safeLocation", Object.class, firstHandle));

      WorkbookAnalysis.AnalysisFinding sheetFinding =
          invoke(
              controller,
              "finding",
              WorkbookAnalysis.AnalysisFinding.class,
              AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
              AnalysisSeverity.ERROR,
              firstHandle,
              "Broken Pivot",
              "Location missing",
              List.of("evidence"));
      assertInstanceOf(WorkbookAnalysis.AnalysisLocation.Sheet.class, sheetFinding.location());
      assertTrue(
          invoke(controller, "resolvedName", String.class, firstHandle)
              .startsWith("_GG_PIVOT_Report_"));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotCacheDefinition cacheDefinition = pivot.getPivotCacheDefinition();
      XSSFPivotCacheRecords cacheRecords =
          cacheDefinition.getRelations().stream()
              .filter(XSSFPivotCacheRecords.class::isInstance)
              .map(XSSFPivotCacheRecords.class::cast)
              .findFirst()
              .orElseThrow();

      assertTrue(PoiRelationRemoval.defaultRemover().test(cacheDefinition, cacheRecords));
      List<AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(codes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_RECORDS));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotCacheDefinition cacheDefinition = pivot.getPivotCacheDefinition();
      invokeVoid(
          controller,
          "removeWorkbookPivotCacheRegistration",
          workbook.xssfWorkbook(),
          pivot.getCTPivotTableDefinition().getCacheId(),
          workbook.xssfWorkbook().getRelationId(cacheDefinition));

      List<AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(codes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_WORKBOOK_CACHE));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      ThrowingPivotTable pivot =
          new ThrowingPivotTable(pivotTableDefinition("Broken Cache Pivot", "C5:F9", 99L));
      Object handle =
          newPivotHandle(
              workbook.xssfWorkbook().getSheetIndex("Report"),
              0,
              "Report",
              workbook.xssfWorkbook().getSheet("Report"),
              pivot);

      assertNull(invoke(controller, "cacheDefinition", XSSFPivotCacheDefinition.class, pivot));
      IllegalArgumentException failure =
          assertInvocationFailure(
              IllegalArgumentException.class,
              () ->
                  invoke(
                      controller,
                      "requiredCacheDefinition",
                      XSSFPivotCacheDefinition.class,
                      pivot));
      assertTrue(failure.getMessage().contains("missing its cache definition relation"));

      @SuppressWarnings("unchecked")
      List<AnalysisFindingCode> codes =
          ((List<WorkbookAnalysis.AnalysisFinding>)
                  invoke(
                      controller,
                      "pivotTableHealthFindings",
                      List.class,
                      workbook.xssfWorkbook(),
                      handle))
              .stream().map(WorkbookAnalysis.AnalysisFinding::code).toList();
      assertTrue(codes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_DEFINITION));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      pivot.getCTPivotTableDefinition().unsetDataFields();

      ExcelPivotTableSnapshot.Unsupported snapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Unsupported.class,
              controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst());
      assertTrue(snapshot.detail().contains("does not contain any data fields"));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      pivot.getCTPivotTableDefinition().getLocation().setRef("not-a-range");

      ExcelPivotTableSnapshot.Unsupported snapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Unsupported.class,
              controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst());
      assertTrue(snapshot.detail().contains("location range"));
      List<AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(codes.contains(AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .unsetWorksheetSource();

      IllegalArgumentException failure =
          assertInvocationFailure(
              IllegalArgumentException.class,
              () ->
                  invoke(
                      controller,
                      "snapshotSource",
                      ExcelPivotTableSnapshot.Source.class,
                      workbook.xssfWorkbook(),
                      pivot));
      assertTrue(failure.getMessage().contains("worksheetSource"));

      List<AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(codes.contains(AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE));
    }
  }
}
