package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Residual pivot-table utility and deletion coverage. */
class ExcelPivotTableResidualCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void directPivotControllerDelegatorsCoverLifecycleAndSnapshotBridges() throws Exception {
    ExcelPivotTableController permissiveController =
        new ExcelPivotTableController((parent, child) -> true);
    assertTrue(
        invoke(
            permissiveController,
            "removePoiRelation",
            Boolean.class,
            new NullDocumentPart(),
            new NullDocumentPart()));

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      PivotHandle handle = (PivotHandle) allPivotHandles(workbook).getFirst();
      ExcelPivotTableSnapshot.Anchor anchor = new ExcelPivotTableSnapshot.Anchor("C5", "C5:F9");
      ExcelPivotTableSnapshot.Unsupported unsupported =
          invoke(
              controller,
              "unsupportedSnapshot",
              ExcelPivotTableSnapshot.Unsupported.class,
              handle,
              "Broken Pivot",
              anchor,
              "missing cache relation");
      assertEquals("Broken Pivot", unsupported.name());
      assertEquals("missing cache relation", unsupported.detail());
      assertEquals(anchor, unsupported.anchor());

      assertEquals(
          List.of("Region", "Stage", "Owner", "Amount"),
          invoke(controller, "cacheFieldNames", List.class, pivot));
      ExcelPivotTableSnapshot.Field sourceField =
          invoke(
              controller,
              "sourceField",
              ExcelPivotTableSnapshot.Field.class,
              List.of("Region", "Amount"),
              1);
      assertEquals("Amount", sourceField.sourceColumnName());

      XSSFPivotCacheDefinition cacheDefinition =
          invoke(controller, "requiredCacheDefinition", XSSFPivotCacheDefinition.class, pivot);
      assertSame(pivot.getPivotCacheDefinition(), cacheDefinition);
      XSSFPivotCacheRecords cacheRecords =
          invoke(controller, "cacheRecords", XSSFPivotCacheRecords.class, cacheDefinition);
      assertNotNull(cacheRecords);
    }
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void pivotControllerCoversResidualDeletionAndUtilityBranches() throws Exception {
    try (ExcelWorkbook workbook = pivotWorkbook()) {
      IllegalArgumentException missingPivot =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deletePivotTable(workbook, "Missing Pivot", "Report"));
      assertTrue(missingPivot.getMessage().contains("pivot table not found"));

      IllegalArgumentException wrongSheet =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deletePivotTable(workbook, "Sales Pivot", "SecondReport"));
      assertTrue(wrongSheet.getMessage().contains("pivot table not found"));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      pivot.setPivotCacheDefinition(null);
      ExcelPivotTableController permissiveController =
          new ExcelPivotTableController((parent, child) -> true);

      invokeVoid(
          permissiveController,
          "deletePivotHandle",
          workbook,
          allPivotHandles(workbook).getFirst());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var pivotCaches = workbook.getCTWorkbook().addNewPivotCaches();
      var pivotCache = pivotCaches.addNewPivotCache();
      pivotCache.setCacheId(11L);
      pivotCache.setId("rIdMatch");

      invokeVoid(controller, "removeWorkbookPivotCacheRegistration", workbook, -1L, null);
      assertTrue(workbook.getCTWorkbook().isSetPivotCaches());
      assertEquals(1, workbook.getCTWorkbook().getPivotCaches().sizeOfPivotCacheArray());

      invokeVoid(controller, "removeWorkbookPivotCacheRegistration", workbook, -1L, "rIdMatch");

      assertFalse(workbook.getCTWorkbook().isSetPivotCaches());
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotTable secondPivot = pivotWorkbookWithSecondPivot(workbook);
      secondPivot.setPivotCacheDefinition(null);

      assertTrue(
          invoke(controller, "pivotTableIdHighWaterMark", Integer.class, workbook.xssfWorkbook())
              >= pivot.getCTPivotTableDefinition().getCacheId());
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      assertFalse(
          invoke(
              controller,
              "cacheDefinitionShared",
              Boolean.class,
              workbook,
              allPivotHandles(workbook).getFirst(),
              new XSSFPivotCacheDefinition()));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Report");
      ExcelPivotTableController permissiveController =
          new ExcelPivotTableController((parent, child) -> true);

      invokeVoid(
          permissiveController,
          "deletePivotHandle",
          workbook,
          newPivotHandle(
              workbook.xssfWorkbook().getSheetIndex("Report"),
              0,
              "Report",
              workbook.xssfWorkbook().getSheet("Report"),
              new SyntheticPivotTable(
                  pivotTableDefinition("Synthetic", "C5:F9", 11L),
                  new XSSFPivotCacheDefinition())));
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

      assertTrue(
          invoke(controller, "pivotTableIdHighWaterMark", Integer.class, workbook.xssfWorkbook())
              >= pivot.getCTPivotTableDefinition().getCacheId());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      PackagePart wrongSuffixPart =
          workbook
              .getPackagePart()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/pivotTables/pivotTable7.bin"),
                  "application/octet-stream");
      assertEquals(
          0,
          invoke(
              controller,
              "packagePartIndex",
              Integer.class,
              wrongSuffixPart,
              "/xl/pivotTables/pivotTable"));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      List<Object> handles = allPivotHandles(workbook);
      assertEquals(
          List.of(handles.getFirst()),
          invoke(
              controller,
              "selectHandlesByName",
              List.class,
              handles,
              List.of("Sales Pivot", "Missing Pivot")));
      assertEquals(
          List.of(),
          invoke(
              controller,
              "snapshotFields",
              List.class,
              new org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField[0],
              List.of("Region")));
      assertEquals(
          List.of(),
          invoke(
              controller,
              "snapshotPageFields",
              List.class,
              new org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField[0],
              List.of("Region")));

      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      var worksheetSource =
          pivot
              .getPivotCacheDefinition()
              .getCTPivotCacheDefinition()
              .getCacheSource()
              .getWorksheetSource();
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "SalesTableA",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      worksheetSource.setRef(" ");
      worksheetSource.setName("SalesTableA");
      assertDoesNotThrow(
          () ->
              invoke(
                  controller,
                  "snapshotSource",
                  ExcelPivotTableSnapshot.Source.class,
                  workbook.xssfWorkbook(),
                  pivot));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .unsetWorksheetSource();

      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "snapshotSource",
                  ExcelPivotTableSnapshot.Source.class,
                  workbook.xssfWorkbook(),
                  pivot));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      removeChild(pivot.getPivotCacheDefinition().getCTPivotCacheDefinition(), "cacheSource");

      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "snapshotSource",
                  ExcelPivotTableSnapshot.Source.class,
                  workbook.xssfWorkbook(),
                  pivot));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      Name workbookScoped = workbook.xssfWorkbook().createName();
      workbookScoped.setNameName("WorkbookScoped");
      workbookScoped.setRefersToFormula("A1:B2");
      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "sourceSheetName",
                  String.class,
                  new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
                  workbookScoped,
                  " "));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      var sheet = workbook.getOrCreateSheet("Data").xssfSheet();
      sheet.createRow(0).createCell(0).setCellValue(42d);
      sheet.getRow(0).createCell(1).setCellValue("Amount");
      sheet.createRow(1).createCell(0).setCellValue("North");
      sheet.getRow(1).createCell(1).setCellValue(10d);

      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "sourceColumns",
                  Object.class,
                  sheet,
                  new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
                  "Data!A1:B2"));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      var sheet = workbook.getOrCreateSheet("Data").xssfSheet();
      sheet.createRow(0).createCell(0).setCellValue("");
      sheet.getRow(0).createCell(1).setCellValue("Amount");
      sheet.createRow(1).createCell(0).setCellValue("North");
      sheet.getRow(1).createCell(1).setCellValue(10d);

      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "sourceColumns",
                  Object.class,
                  sheet,
                  new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
                  "Data!A1:B2"));
    }

    try (ExcelWorkbook workbook = pivotWorkbook()) {
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "sourceColumnName", String.class, List.of("Only"), -1));

      short blankFormatId = workbook.xssfWorkbook().createDataFormat().getFormat(" ");
      assertNull(
          invoke(
              controller,
              "numberFormat",
              String.class,
              workbook.xssfWorkbook(),
              (long) blankFormatId));

      Object handle = allPivotHandles(workbook).getFirst();
      workbook.xssfWorkbook().getPivotTables().getFirst().getCTPivotTableDefinition().setName(" ");
      assertNull(invoke(controller, "actualName", String.class, handle));

      assertEquals("Pivot_1", invoke(controller, "sanitize", String.class, "Pivot_1"));
      assertEquals(
          "A1:B1",
          invoke(
              controller,
              "normalizeArea",
              String.class,
              new AreaReference("A1:B1", SpreadsheetVersion.EXCEL2007)));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "requireNonBlank", String.class, null, "message"));
    }
  }
}
