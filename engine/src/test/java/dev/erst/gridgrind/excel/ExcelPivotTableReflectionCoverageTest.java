package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCache;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheRecords;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/** Pivot-table reflective and helper-branch coverage. */
class ExcelPivotTableReflectionCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void reflectiveHelpersCoverRemainingPivotUtilityBranches() throws Exception {
    try (ExcelWorkbook workbook = pivotWorkbook()) {
      XSSFPivotTable pivot = workbook.xssfWorkbook().getPivotTables().getFirst();
      XSSFPivotCacheDefinition cacheDefinition = pivot.getPivotCacheDefinition();
      long originalCacheId = pivot.getCTPivotTableDefinition().getCacheId();
      XSSFPivotTable secondPivot = pivotWorkbookWithSecondPivot(workbook);
      long duplicateCacheId = pivot.getPivotCache().getCTPivotCache().getCacheId();

      invokeVoid(controller, "normalizeCacheId", workbook.xssfWorkbook(), pivot);
      assertEquals(originalCacheId, pivot.getCTPivotTableDefinition().getCacheId());
      secondPivot.getPivotCache().getCTPivotCache().setCacheId(duplicateCacheId);
      secondPivot.getCTPivotTableDefinition().setCacheId(duplicateCacheId);
      invokeVoid(controller, "normalizeCacheId", workbook.xssfWorkbook(), secondPivot);
      assertNotEquals(duplicateCacheId, secondPivot.getCTPivotTableDefinition().getCacheId());
      invokeVoid(
          controller,
          "normalizeCacheId",
          workbook.xssfWorkbook(),
          new NoCachePivotTable(pivotTableDefinition("No Cache Pivot", "C5:F9", 17L)));
      CacheOnlyPivotTable cacheOnlyPivot =
          new CacheOnlyPivotTable(
              pivotTableDefinition("Cache Only Pivot", "C5:F9", 17L),
              new XSSFPivotCache(CTPivotCache.Factory.newInstance()));
      cacheOnlyPivot.getPivotCache().getCTPivotCache().setCacheId(duplicateCacheId);
      invokeVoid(controller, "normalizeCacheId", workbook.xssfWorkbook(), cacheOnlyPivot);
      assertNotEquals(duplicateCacheId, cacheOnlyPivot.getCTPivotTableDefinition().getCacheId());

      Object sourceColumns =
          invoke(
              controller,
              "sourceColumns",
              Object.class,
              workbook.xssfWorkbook().getSheet("Data"),
              new AreaReference("A1:D5", SpreadsheetVersion.EXCEL2007),
              "Data!A1:D5");
      assertEquals(0, invoke(sourceColumns, "relativeIndex", Integer.class, "Region"));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(sourceColumns, "relativeIndex", Integer.class, "Missing"));

      assertEquals(
          "Amount",
          invoke(controller, "sourceColumnName", String.class, List.of("Region", "Amount"), 1));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "sourceColumnName", String.class, List.of("Only"), 1));
      assertEquals(
          ExcelPivotDataConsolidateFunction.SUM,
          invoke(
              controller,
              "fromSubtotal",
              ExcelPivotDataConsolidateFunction.class,
              DataConsolidateFunction.SUM.getValue()));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "fromSubtotal", ExcelPivotDataConsolidateFunction.class, -999));

      assertNull(invoke(controller, "numberFormat", String.class, workbook.xssfWorkbook(), null));
      assertEquals(
          "General", invoke(controller, "numberFormat", String.class, workbook.xssfWorkbook(), 0L));
      assertNull(
          invoke(
              controller, "numberFormat", String.class, workbook.xssfWorkbook(), Long.MAX_VALUE));

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "ScopedRange",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "ScopedRange",
              new ExcelNamedRangeScope.SheetScope("Data"),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      assertEquals(
          2,
          invoke(
                  controller,
                  "matchingNamedRanges",
                  List.class,
                  workbook.xssfWorkbook(),
                  "ScopedRange",
                  null)
              .size());
      assertEquals(
          2,
          invoke(
                  controller,
                  "matchingNamedRanges",
                  List.class,
                  workbook.xssfWorkbook(),
                  "ScopedRange",
                  "Data")
              .size());
      assertEquals(
          1,
          invoke(
                  controller,
                  "matchingNamedRanges",
                  List.class,
                  workbook.xssfWorkbook(),
                  "ScopedRange",
                  "Other")
              .size());

      Name blankName = workbook.xssfWorkbook().createName();
      blankName.setNameName("BlankSource");
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "namedRangeArea", AreaReference.class, blankName));

      Name invalidName = workbook.xssfWorkbook().createName();
      invalidName.setNameName("InvalidSource");
      invalidName.setRefersToFormula("A1,B2");
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "namedRangeArea", AreaReference.class, invalidName));

      Name sheetScoped = workbook.xssfWorkbook().createName();
      sheetScoped.setNameName("SheetScoped");
      sheetScoped.setSheetIndex(workbook.xssfWorkbook().getSheetIndex("Data"));
      sheetScoped.setRefersToFormula("A1:B2");
      Name workbookScoped = workbook.xssfWorkbook().createName();
      workbookScoped.setNameName("WorkbookScoped");
      workbookScoped.setRefersToFormula("A1:B2");
      assertEquals(
          "Data",
          invoke(
              controller,
              "sourceSheetName",
              String.class,
              new AreaReference("Data!$A$1:$B$2", SpreadsheetVersion.EXCEL2007),
              sheetScoped,
              null));
      assertEquals(
          "Data",
          invoke(
              controller,
              "sourceSheetName",
              String.class,
              new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
              sheetScoped,
              null));
      assertEquals(
          "Fallback",
          invoke(
              controller,
              "sourceSheetName",
              String.class,
              new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
              workbookScoped,
              "Fallback"));
      assertEquals(
          "Fallback",
          invoke(
              controller,
              "sourceSheetName",
              String.class,
              new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
              null,
              "Fallback"));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "sourceSheetName",
                  String.class,
                  new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007),
                  workbookScoped,
                  null));

      populatePivotSource(workbook, "TableDataA");
      populatePivotSource(workbook, "TableDataB");
      workbook.setTable(
          new ExcelTableDefinition(
              "SalesTableA", "TableDataA", "A1:D5", false, new ExcelTableStyle.None()));
      workbook.setTable(
          new ExcelTableDefinition(
              "SalesTableB", "TableDataB", "A1:D5", false, new ExcelTableStyle.None()));
      XSSFTable duplicateTable =
          workbook.xssfWorkbook().getSheet("TableDataB").getTables().getFirst();
      duplicateTable.setName("SalesTableA");
      duplicateTable.setDisplayName("SalesTableA");
      populatePivotSource(workbook, "TableDataC");
      workbook.setTable(
          new ExcelTableDefinition(
              "SheetTableA", "TableDataC", "A1:D5", false, new ExcelTableStyle.None()));
      workbook
          .getOrCreateSheet("TableDataC")
          .setRange(
              "F1:I5",
              List.of(
                  List.of(
                      ExcelCellValue.text("Region"),
                      ExcelCellValue.text("Stage"),
                      ExcelCellValue.text("Owner"),
                      ExcelCellValue.text("Amount")),
                  List.of(
                      ExcelCellValue.text("North"),
                      ExcelCellValue.text("Plan"),
                      ExcelCellValue.text("Ada"),
                      ExcelCellValue.number(10)),
                  List.of(
                      ExcelCellValue.text("North"),
                      ExcelCellValue.text("Do"),
                      ExcelCellValue.text("Ada"),
                      ExcelCellValue.number(15)),
                  List.of(
                      ExcelCellValue.text("South"),
                      ExcelCellValue.text("Plan"),
                      ExcelCellValue.text("Lin"),
                      ExcelCellValue.number(7)),
                  List.of(
                      ExcelCellValue.text("South"),
                      ExcelCellValue.text("Do"),
                      ExcelCellValue.text("Lin"),
                      ExcelCellValue.number(12))));
      workbook.setTable(
          new ExcelTableDefinition(
              "SheetTableB", "TableDataC", "F1:I5", false, new ExcelTableStyle.None()));
      XSSFTable sameSheetDuplicate =
          workbook.xssfWorkbook().getSheet("TableDataC").getTables().get(1);
      sameSheetDuplicate.setName("SheetTableA");
      sameSheetDuplicate.setDisplayName("SheetTableA");

      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "tableByName",
                  XSSFTable.class,
                  workbook.xssfWorkbook(),
                  "SalesTableA",
                  null));
      assertNotNull(
          invoke(
              controller,
              "tableByName",
              XSSFTable.class,
              workbook.xssfWorkbook(),
              "SalesTableA",
              "TableDataA"));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(controller, "requiredTableByName", XSSFTable.class, workbook, "MissingTable"));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "tableByName",
                  XSSFTable.class,
                  workbook.xssfWorkbook(),
                  "SheetTableA",
                  "TableDataC"));

      CTPivotTableDefinition syntheticDefinition = CTPivotTableDefinition.Factory.newInstance();
      syntheticDefinition.addNewColFields().addNewField().setX(-2);
      syntheticDefinition.getColFields().addNewField().setX(1);
      var syntheticDataFields = syntheticDefinition.addNewDataFields();
      var syntheticDataField = syntheticDataFields.addNewDataField();
      syntheticDataField.setFld(1);
      syntheticDataField.setName("");
      syntheticDataField.xsetSubtotal(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction.Factory
              .newValue("max"));
      assertTrue(syntheticDataField.isSetSubtotal());
      syntheticDataField.setNumFmtId(Long.MAX_VALUE);
      Object columnAxisSnapshot =
          invoke(
              controller,
              "snapshotColumnLabels",
              Object.class,
              syntheticDefinition,
              List.of("Region", "Amount"));
      @SuppressWarnings("unchecked")
      List<ExcelPivotTableSnapshot.Field> columnLabels =
          (List<ExcelPivotTableSnapshot.Field>)
              invoke(columnAxisSnapshot, "columnLabels", List.class);
      assertEquals(List.of("Amount"), fieldNames(columnLabels));
      assertTrue(invoke(columnAxisSnapshot, "valuesAxisOnColumns", Boolean.class));

      assertEquals(
          List.of(),
          invoke(controller, "snapshotFields", List.class, null, List.of("Region", "Amount")));
      assertEquals(
          List.of(),
          invoke(
              controller,
              "snapshotDataFields",
              List.class,
              workbook.xssfWorkbook(),
              (CTPivotTableDefinition)
                  java.lang.reflect.Proxy.newProxyInstance(
                      Thread.currentThread().getContextClassLoader(),
                      new Class<?>[] {CTPivotTableDefinition.class},
                      (proxy, method, args) -> {
                        if ("getDataFields".equals(method.getName())) {
                          return null;
                        }
                        throw new UnsupportedOperationException(method.getName());
                      }),
              List.of("Region", "Amount")));
      CTPivotTableDefinition emptyDataFieldsDefinition =
          CTPivotTableDefinition.Factory.newInstance();
      emptyDataFieldsDefinition.addNewDataFields();
      assertEquals(
          List.of(),
          invoke(
              controller,
              "snapshotDataFields",
              List.class,
              workbook.xssfWorkbook(),
              emptyDataFieldsDefinition,
              List.of("Region", "Amount")));
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField nullSubtotalField =
          (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField)
              java.lang.reflect.Proxy.newProxyInstance(
                  Thread.currentThread().getContextClassLoader(),
                  new Class<?>[] {
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField.class
                  },
                  (proxy, method, args) ->
                      switch (method.getName()) {
                        case "getFld" -> 1L;
                        case "getSubtotal" -> null;
                        case "getName" -> null;
                        case "isSetNumFmtId" -> false;
                        default -> throw new UnsupportedOperationException(method.getName());
                      });
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields nullSubtotalFields =
          (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields)
              java.lang.reflect.Proxy.newProxyInstance(
                  Thread.currentThread().getContextClassLoader(),
                  new Class<?>[] {
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields.class
                  },
                  (proxy, method, args) ->
                      switch (method.getName()) {
                        case "sizeOfDataFieldArray" -> 1;
                        case "getDataFieldArray" ->
                            new org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField[] {
                              nullSubtotalField
                            };
                        default -> throw new UnsupportedOperationException(method.getName());
                      });
      CTPivotTableDefinition defaultSubtotalDefinition =
          (CTPivotTableDefinition)
              java.lang.reflect.Proxy.newProxyInstance(
                  Thread.currentThread().getContextClassLoader(),
                  new Class<?>[] {CTPivotTableDefinition.class},
                  (proxy, method, args) -> {
                    if ("getDataFields".equals(method.getName())) {
                      return nullSubtotalFields;
                    }
                    throw new UnsupportedOperationException(method.getName());
                  });
      @SuppressWarnings("unchecked")
      List<ExcelPivotTableSnapshot.DataField> defaultSubtotalSnapshots =
          (List<ExcelPivotTableSnapshot.DataField>)
              invoke(
                  controller,
                  "snapshotDataFields",
                  List.class,
                  workbook.xssfWorkbook(),
                  defaultSubtotalDefinition,
                  List.of("Region", "Amount"));
      assertEquals(
          ExcelPivotDataConsolidateFunction.SUM, defaultSubtotalSnapshots.getFirst().function());
      @SuppressWarnings("unchecked")
      List<ExcelPivotTableSnapshot.DataField> syntheticDataSnapshots =
          (List<ExcelPivotTableSnapshot.DataField>)
              invoke(
                  controller,
                  "snapshotDataFields",
                  List.class,
                  workbook.xssfWorkbook(),
                  syntheticDefinition,
                  List.of("Region", "Amount"));
      assertEquals("Amount", syntheticDataSnapshots.getFirst().displayName());
      assertNull(syntheticDataSnapshots.getFirst().valueFormat());
      assertEquals(
          ExcelPivotDataConsolidateFunction.MAX, syntheticDataSnapshots.getFirst().function());

      ExcelPivotTableSnapshot.Unsupported unsupportedSnapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Unsupported.class,
              invoke(
                  controller,
                  "snapshot",
                  ExcelPivotTableSnapshot.class,
                  workbook.xssfWorkbook(),
                  newPivotHandle(
                      workbook.xssfWorkbook().getSheetIndex("Report"),
                      0,
                      "Report",
                      workbook.xssfWorkbook().getSheet("Report"),
                      new ThrowingPivotTable(
                          pivotTableDefinition("Broken Snapshot", "C5:F9", 42L)))));
      assertTrue(unsupportedSnapshot.detail().contains("missing its cache definition relation"));
      ExcelPivotTableSnapshot.Unsupported missingLocationSnapshot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Unsupported.class,
              invoke(
                  controller,
                  "snapshot",
                  ExcelPivotTableSnapshot.class,
                  workbook.xssfWorkbook(),
                  newPivotHandle(
                      workbook.xssfWorkbook().getSheetIndex("Report"),
                      0,
                      "Report",
                      workbook.xssfWorkbook().getSheet("Report"),
                      new ThrowingPivotTable(CTPivotTableDefinition.Factory.newInstance()))));
      assertTrue(missingLocationSnapshot.detail().contains("location range"));

      assertNotNull(
          invoke(
              controller,
              "snapshotSource",
              ExcelPivotTableSnapshot.Source.class,
              workbook.xssfWorkbook(),
              pivot));
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .getWorksheetSource()
          .setSheet(" ");
      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "snapshotSource",
                  ExcelPivotTableSnapshot.Source.class,
                  workbook.xssfWorkbook(),
                  pivot));
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .getWorksheetSource()
          .setSheet("Data");
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .getWorksheetSource()
          .unsetRef();
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .getWorksheetSource()
          .setName(" ");
      assertInvocationFailure(
          IllegalArgumentException.class,
          () ->
              invoke(
                  controller,
                  "snapshotSource",
                  ExcelPivotTableSnapshot.Source.class,
                  workbook.xssfWorkbook(),
                  pivot));
      pivot
          .getPivotCacheDefinition()
          .getCTPivotCacheDefinition()
          .getCacheSource()
          .getWorksheetSource()
          .setName("SalesTableA");

      removeChild(cacheDefinition.getCTPivotCacheDefinition(), "cacheFields");
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "cacheFieldNames", List.class, pivot));

      assertEquals("Pivot_Name___", invoke(controller, "sanitize", String.class, "Pivot Name!?/"));
      assertEquals(
          "A1",
          invoke(
              controller,
              "normalizeArea",
              String.class,
              new AreaReference("A1", SpreadsheetVersion.EXCEL2007)));
      assertEquals(
          "A1:B2",
          invoke(
              controller,
              "normalizeArea",
              String.class,
              new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007)));
      assertEquals(
          "value", invoke(controller, "requireNonBlank", String.class, "value", "message"));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "requireNonBlank", String.class, " ", "message"));
      assertEquals(
          "fallback", invoke(controller, "nonBlankOrDefault", String.class, " ", "fallback"));
      assertEquals(
          "value", invoke(controller, "nonBlankOrDefault", String.class, "value", "fallback"));
      assertInvocationFailure(
          IllegalArgumentException.class,
          () -> invoke(controller, "contiguousArea", AreaReference.class, "A1,B2", "range"));

      assertNotNull(
          invoke(
              controller,
              "firstRelation",
              Object.class,
              secondPivot.getPivotCacheDefinition(),
              XSSFPivotCacheRecords.class));
      assertNull(
          invoke(
              controller,
              "firstRelation",
              Object.class,
              secondPivot.getPivotCacheDefinition(),
              XSSFPivotTable.class));
      assertNull(
          invoke(
              controller,
              "firstRelation",
              Object.class,
              new NullDocumentPart(),
              XSSFPivotCacheRecords.class));
      assertNull(
          invoke(controller, "workbookPivotCache", Object.class, workbook.xssfWorkbook(), 9999L));

      assertEquals(
          0,
          invoke(
              controller,
              "packagePartIndex",
              Integer.class,
              new NullDocumentPart(),
              "/xl/pivotTables/pivotTable"));
      assertEquals(
          0,
          invoke(
              controller,
              "packagePartIndex",
              Integer.class,
              workbook.xssfWorkbook(),
              "/xl/pivotTables/pivotTable"));
      PackagePart syntheticPart =
          workbook
              .xssfWorkbook()
              .getPackagePart()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/pivotTables/pivotTableX.xml"),
                  "application/xml");
      syntheticPart.addExternalRelationship("https://example.com/object", "external");
      assertEquals(
          0,
          invoke(
              controller,
              "packagePartIndex",
              Integer.class,
              new SyntheticDocumentPart(syntheticPart),
              "/xl/pivotTables/pivotTable"));
      assertTrue(
          invoke(controller, "pivotTableIdHighWaterMark", Integer.class, workbook.xssfWorkbook())
              >= secondPivot.getCTPivotTableDefinition().getCacheId());

      PackagePart orphanPart =
          workbook
              .xssfWorkbook()
              .getPackagePart()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/pivotCache/pivotCacheRecords99.xml"),
                  "application/xml");
      workbook
          .xssfWorkbook()
          .getPackagePart()
          .getPackage()
          .createPart(
              PackagingURIHelper.createPartName("/xl/_rels/fakePivot.rels"),
              "application/vnd.openxmlformats-package.relationships+xml");
      assertEquals(
          99,
          invoke(controller, "pivotTableIdHighWaterMark", Integer.class, workbook.xssfWorkbook()));
      invokeVoid(
          controller, "cleanupPackagePartIfUnused", workbook.xssfWorkbook().getPackage(), null);
      invokeVoid(
          controller,
          "cleanupPackagePartIfUnused",
          workbook.xssfWorkbook().getPackage(),
          secondPivot.getPivotCacheDefinition().getPackagePart().getPartName());
      assertTrue(
          workbook
              .xssfWorkbook()
              .getPackage()
              .containPart(secondPivot.getPivotCacheDefinition().getPackagePart().getPartName()));
      invokeVoid(
          controller,
          "cleanupPackagePartIfUnused",
          workbook.xssfWorkbook().getPackage(),
          orphanPart.getPartName());
      assertFalse(workbook.xssfWorkbook().getPackage().containPart(orphanPart.getPartName()));
      invokeVoid(
          controller,
          "cleanupPackagePartIfUnused",
          workbook.xssfWorkbook().getPackage(),
          orphanPart.getPartName());

      try (XSSFWorkbook emptyWorkbook = new XSSFWorkbook()) {
        invokeVoid(controller, "removeWorkbookPivotCacheRegistration", emptyWorkbook, 1L, null);
      }
      int pivotCacheCountBefore =
          workbook.xssfWorkbook().getCTWorkbook().getPivotCaches().sizeOfPivotCacheArray();
      invokeVoid(
          controller,
          "removeWorkbookPivotCacheRegistration",
          workbook.xssfWorkbook(),
          -1L,
          "missing-relation");
      assertEquals(
          pivotCacheCountBefore,
          workbook.xssfWorkbook().getCTWorkbook().getPivotCaches().sizeOfPivotCacheArray());

      List<Object> handles = allPivotHandles(workbook);
      assertFalse(
          invoke(
              controller,
              "cacheDefinitionShared",
              Boolean.class,
              workbook,
              handles.getFirst(),
              pivot.getPivotCacheDefinition()));
      assertTrue(
          invoke(
              controller,
              "cacheDefinitionShared",
              Boolean.class,
              workbook,
              handles.getFirst(),
              secondPivot.getPivotCacheDefinition()));

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
              new ThrowingPivotTable(pivotTableDefinition("Deleted Broken Pivot", "C5:F9", 9L))));
      assertNull(
          invoke(
              controller,
              "rawLocationRange",
              Object.class,
              newPivotHandle(
                  workbook.xssfWorkbook().getSheetIndex("Report"),
                  0,
                  "Report",
                  workbook.xssfWorkbook().getSheet("Report"),
                  new ThrowingPivotTable(CTPivotTableDefinition.Factory.newInstance()))));
    }

    try (XSSFWorkbook orphanedWorkbook = new XSSFWorkbook()) {
      orphanedWorkbook.getCTWorkbook().addNewPivotCaches().addNewPivotCache().setCacheId(1L);
      IllegalStateException failure =
          assertInvocationFailure(
              IllegalStateException.class,
              () -> invokeVoid(controller, "primePivotTableAllocator", orphanedWorkbook, null));
      assertTrue(failure.getMessage().contains("without any live pivot relations"));
    }
  }
}
