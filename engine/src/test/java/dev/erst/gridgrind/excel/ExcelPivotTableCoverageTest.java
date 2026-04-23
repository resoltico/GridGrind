package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.ooxml.POIXMLDocumentPart;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/** Covers reflective and malformed-state branches for the Phase 7 pivot-table engine surface. */
@SuppressWarnings({"PMD.CommentRequired", "PMD.NcssCount", "PMD.SignatureDeclareThrowsException"})
class ExcelPivotTableCoverageTest {
  private final ExcelPivotTableController controller = new ExcelPivotTableController();

  @Test
  void setPivotTableRejectsInvalidAuthoringInputsAndReplacesSameSheetPivot() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      workbook.getOrCreateSheet("OtherReport");

      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "C5",
              List.of(),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "E5",
              List.of("Owner"),
              List.of("Stage"),
              List.of()));

      ExcelPivotTableSnapshot.Supported replaced =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Supported.class,
              controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst());
      assertEquals(1, controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).size());
      assertEquals("E5", replaced.anchor().topLeftAddress());
      assertEquals(List.of("Stage"), fieldNames(replaced.rowLabels()));
      assertEquals(List.of("Owner"), fieldNames(replaced.reportFilters()));

      IllegalArgumentException differentSheetFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Sales Pivot",
                          "OtherReport",
                          new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(differentSheetFailure.getMessage().contains("different sheet"));

      assertThrows(
          SheetNotFoundException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Report",
                      "MissingReport",
                      new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Column",
                      "Report",
                      new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                      "C5",
                      List.of(),
                      List.of("Missing"),
                      List.of())));

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "AmbiguousSource",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "AmbiguousSource",
              new ExcelNamedRangeScope.SheetScope("Data"),
              new ExcelNamedRangeTarget("Data", "A1:D5")));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Named Range",
                      "Report",
                      new ExcelPivotTableDefinition.Source.NamedRange("MissingSource"),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));
      IllegalArgumentException ambiguousNamedRangeFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Ambiguous Named Range",
                          "Report",
                          new ExcelPivotTableDefinition.Source.NamedRange("AmbiguousSource"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(ambiguousNamedRangeFailure.getMessage().contains("ambiguous"));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Table",
                      "Report",
                      new ExcelPivotTableDefinition.Source.Table("MissingTable"),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));

      workbook
          .getOrCreateSheet("OneRow")
          .setRange(
              "A1:D1",
              List.of(
                  List.of(
                      ExcelCellValue.text("Region"),
                      ExcelCellValue.text("Stage"),
                      ExcelCellValue.text("Owner"),
                      ExcelCellValue.text("Amount"))));
      IllegalArgumentException oneRowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "One Row",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("OneRow", "A1:D1"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(oneRowFailure.getMessage().contains("header row plus at least one data row"));

      ExcelSheet missingHeaderSheet = workbook.getOrCreateSheet("MissingHeader");
      missingHeaderSheet.setCell("A3", ExcelCellValue.text("North"));
      missingHeaderSheet.setCell("B3", ExcelCellValue.number(10));
      IllegalArgumentException missingHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Missing Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("MissingHeader", "A2:B3"),
                          "C5",
                          List.of(),
                          List.of("North"),
                          List.of())));
      assertTrue(missingHeaderFailure.getMessage().contains("missing its header row"));

      workbook
          .getOrCreateSheet("NumericHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.number(1), ExcelCellValue.text("Amount")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      IllegalArgumentException numericHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Numeric Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("NumericHeader", "A1:B2"),
                          "C5",
                          List.of(),
                          List.of("Amount"),
                          List.of())));
      assertFalse(numericHeaderFailure.getMessage().isBlank());

      workbook
          .getOrCreateSheet("BlankHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.text(" "), ExcelCellValue.text("Amount")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Blank Header",
                      "Report",
                      new ExcelPivotTableDefinition.Source.Range("BlankHeader", "A1:B2"),
                      "C5",
                      List.of(),
                      List.of("Amount"),
                      List.of())));

      ExcelSheet missingHeaderCellSheet = workbook.getOrCreateSheet("MissingHeaderCell");
      missingHeaderCellSheet.setCell("A1", ExcelCellValue.text("Region"));
      missingHeaderCellSheet.setCell("C1", ExcelCellValue.text("Amount"));
      missingHeaderCellSheet.setCell("A2", ExcelCellValue.text("North"));
      missingHeaderCellSheet.setCell("B2", ExcelCellValue.text("Plan"));
      missingHeaderCellSheet.setCell("C2", ExcelCellValue.number(10));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Header Cell",
                      "Report",
                      new ExcelPivotTableDefinition.Source.Range("MissingHeaderCell", "A1:C2"),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));

      workbook
          .getOrCreateSheet("DuplicateHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.text("Region"), ExcelCellValue.text("region")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      IllegalArgumentException duplicateHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Duplicate Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("DuplicateHeader", "A1:B2"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(duplicateHeaderFailure.getMessage().contains("unique case-insensitively"));

      workbook
          .getOrCreateSheet("BlankHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.text(""), ExcelCellValue.text("Amount")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      IllegalArgumentException blankHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Blank Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("BlankHeader", "A1:B2"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(blankHeaderFailure.getMessage().contains("blank header cell"));
    }
  }

  @Test
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

      List<WorkbookAnalysis.AnalysisFindingCode> duplicateCodes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(
          duplicateCodes.contains(WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_DUPLICATE_NAME));

      firstPivot.getCTPivotTableDefinition().setName(null);
      ExcelPivotTableSnapshot syntheticNameSnapshot =
          controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst();
      assertTrue(syntheticNameSnapshot.name().startsWith("_GG_PIVOT_"));

      Object firstHandle = allPivotHandles(workbook).getFirst();
      firstPivot.getCTPivotTableDefinition().getLocation().setRef(null);
      assertNull(invoke(controller, "rawLocationRange", Object.class, firstHandle));
      assertNull(invoke(controller, "safeLocation", Object.class, firstHandle));

      WorkbookAnalysis.AnalysisFinding sheetFinding =
          invoke(
              controller,
              "finding",
              WorkbookAnalysis.AnalysisFinding.class,
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
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
      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(
          codes.contains(WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_RECORDS));
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

      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(
          codes.contains(WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_WORKBOOK_CACHE));
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
      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          ((List<WorkbookAnalysis.AnalysisFinding>)
                  invoke(
                      controller,
                      "pivotTableHealthFindings",
                      List.class,
                      workbook.xssfWorkbook(),
                      handle))
              .stream().map(WorkbookAnalysis.AnalysisFinding::code).toList();
      assertTrue(
          codes.contains(
              WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_DEFINITION));
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
      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(
          codes.contains(WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL));
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

      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.pivotTableHealthFindings(workbook, new ExcelPivotTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE));
    }
  }

  @Test
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

  @Test
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

  private XSSFPivotTable pivotWorkbookWithSecondPivot(ExcelWorkbook workbook) {
    workbook.getOrCreateSheet("SecondReport");
    controller.setPivotTable(
        workbook,
        definition(
            "Second Pivot",
            "SecondReport",
            new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
            "C5",
            List.of(),
            List.of("Stage"),
            List.of()));
    return workbook.xssfWorkbook().getPivotTables().get(1);
  }

  private ExcelWorkbook pivotWorkbook() throws Exception {
    ExcelWorkbook workbook = ExcelWorkbook.create();
    populatePivotSource(workbook, "Data");
    workbook.getOrCreateSheet("Report");
    controller.setPivotTable(
        workbook,
        definition(
            "Sales Pivot",
            "Report",
            new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
            "C5",
            List.of(),
            List.of("Region"),
            List.of()));
    return workbook;
  }

  private ExcelPivotTableDefinition definition(
      String name,
      String sheetName,
      ExcelPivotTableDefinition.Source source,
      String anchor,
      List<String> reportFilters,
      List<String> rowLabels,
      List<String> columnLabels) {
    return new ExcelPivotTableDefinition(
        name,
        sheetName,
        source,
        new ExcelPivotTableDefinition.Anchor(anchor),
        rowLabels,
        columnLabels,
        reportFilters,
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", "#,##0.00")));
  }

  private void populatePivotSource(ExcelWorkbook workbook, String sheetName) {
    workbook
        .getOrCreateSheet(sheetName)
        .setRange(
            "A1:D5",
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
  }

  private List<String> fieldNames(List<ExcelPivotTableSnapshot.Field> fields) {
    return fields.stream().map(ExcelPivotTableSnapshot.Field::sourceColumnName).toList();
  }

  @SuppressWarnings("unchecked")
  private List<Object> allPivotHandles(ExcelWorkbook workbook) throws Exception {
    return (List<Object>) invoke(controller, "allPivotTables", List.class, workbook);
  }

  private Object newPivotHandle(
      int sheetIndex,
      int ordinalOnSheet,
      String sheetName,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      XSSFPivotTable table) {
    return new PivotHandle(sheetIndex, ordinalOnSheet, sheetName, sheet, table);
  }

  private CTPivotTableDefinition pivotTableDefinition(
      String name, String locationRange, long cacheId) {
    CTPivotTableDefinition definition = CTPivotTableDefinition.Factory.newInstance();
    definition.setName(name);
    definition.addNewLocation().setRef(locationRange);
    definition.setCacheId(cacheId);
    return definition;
  }

  private static <T> T invoke(Object target, String name, Class<T> returnType, Object... args)
      throws Exception {
    return returnType.cast(dispatch(target, name, args));
  }

  private static void invokeVoid(Object target, String name, Object... args) throws Exception {
    dispatch(target, name, args);
  }

  private static Object dispatch(Object target, String name, Object... args) throws Exception {
    if (target instanceof ExcelPivotTableController controller) {
      return switch (name) {
        case "actualName" -> ExcelPivotTableIdentitySupport.actualName((PivotHandle) args[0]);
        case "allPivotTables" -> controller.allPivotTables((ExcelWorkbook) args[0]);
        case "cacheDefinition" -> controller.cacheDefinition((XSSFPivotTable) args[0]);
        case "cacheDefinitionShared" ->
            controller.cacheDefinitionShared(
                (ExcelWorkbook) args[0], (PivotHandle) args[1], (XSSFPivotCacheDefinition) args[2]);
        case "cacheFieldNames" -> controller.cacheFieldNames((XSSFPivotTable) args[0]);
        case "cleanupPackagePartIfUnused" -> {
          controller.cleanupPackagePartIfUnused(
              (org.apache.poi.openxml4j.opc.OPCPackage) args[0],
              (org.apache.poi.openxml4j.opc.PackagePartName) args[1]);
          yield null;
        }
        case "contiguousArea" ->
            ExcelPivotTableIdentitySupport.contiguousArea((String) args[0], (String) args[1]);
        case "deletePivotHandle" -> {
          controller.deletePivotHandle((ExcelWorkbook) args[0], (PivotHandle) args[1]);
          yield null;
        }
        case "finding" ->
            controller.finding(
                (WorkbookAnalysis.AnalysisFindingCode) args[0],
                (WorkbookAnalysis.AnalysisSeverity) args[1],
                (PivotHandle) args[2],
                (String) args[3],
                (String) args[4],
                (List<String>) args[5]);
        case "firstRelation" ->
            controller.firstRelation(
                (POIXMLDocumentPart) args[0], (Class<? extends POIXMLDocumentPart>) args[1]);
        case "fromSubtotal" -> controller.fromSubtotal((Integer) args[0]);
        case "matchingNamedRanges" ->
            ExcelPivotTableSourceSupport.matchingNamedRanges(
                (XSSFWorkbook) args[0], (String) args[1], (String) args[2]);
        case "namedRangeArea" -> ExcelPivotTableSourceSupport.namedRangeArea((Name) args[0]);
        case "nonBlankOrDefault" ->
            ExcelPivotTableIdentitySupport.nonBlankOrDefault((String) args[0], (String) args[1]);
        case "normalizeCacheId" -> {
          controller.normalizeCacheId((XSSFWorkbook) args[0], (XSSFPivotTable) args[1]);
          yield null;
        }
        case "normalizeArea" ->
            ExcelPivotTableIdentitySupport.normalizeArea((AreaReference) args[0]);
        case "numberFormat" -> controller.numberFormat((XSSFWorkbook) args[0], (Long) args[1]);
        case "packagePartIndex" -> {
          if (args[0] instanceof PackagePart part) {
            yield controller.packagePartIndex(part, (String) args[1]);
          }
          yield controller.packagePartIndex((POIXMLDocumentPart) args[0], (String) args[1]);
        }
        case "pivotTableHealthFindings" ->
            controller.pivotTableHealthFindings((XSSFWorkbook) args[0], (PivotHandle) args[1]);
        case "pivotTableIdHighWaterMark" ->
            controller.pivotTableIdHighWaterMark((XSSFWorkbook) args[0]);
        case "primePivotTableAllocator" -> {
          controller.primePivotTableAllocator((XSSFWorkbook) args[0], (XSSFPivotTable) args[1]);
          yield null;
        }
        case "rawLocationRange" ->
            ExcelPivotTableIdentitySupport.rawLocationRange((PivotHandle) args[0]);
        case "removeWorkbookPivotCacheRegistration" -> {
          controller.removeWorkbookPivotCacheRegistration(
              (XSSFWorkbook) args[0], (Long) args[1], (String) args[2]);
          yield null;
        }
        case "requireNonBlank" ->
            ExcelPivotTableIdentitySupport.requireNonBlank((String) args[0], (String) args[1]);
        case "requiredCacheDefinition" ->
            controller.requiredCacheDefinition((XSSFPivotTable) args[0]);
        case "requiredTableByName" ->
            ExcelPivotTableSourceSupport.requiredTableByName(
                (ExcelWorkbook) args[0], (String) args[1]);
        case "resolvedName" -> ExcelPivotTableIdentitySupport.resolvedName((PivotHandle) args[0]);
        case "safeLocation" -> ExcelPivotTableIdentitySupport.safeLocation((PivotHandle) args[0]);
        case "sanitize" -> ExcelPivotTableIdentitySupport.sanitize((String) args[0]);
        case "selectHandlesByName" ->
            controller.selectHandlesByName((List<PivotHandle>) args[0], (List<String>) args[1]);
        case "snapshot" -> controller.snapshot((XSSFWorkbook) args[0], (PivotHandle) args[1]);
        case "snapshotColumnLabels" ->
            controller.snapshotColumnLabels(
                (CTPivotTableDefinition) args[0], (List<String>) args[1]);
        case "snapshotDataFields" ->
            controller.snapshotDataFields(
                (XSSFWorkbook) args[0], (CTPivotTableDefinition) args[1], (List<String>) args[2]);
        case "snapshotFields" ->
            controller.snapshotFields(
                args[0] == null ? null : (CTField[]) args[0], (List<String>) args[1]);
        case "snapshotPageFields" ->
            controller.snapshotPageFields(
                (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField[]) args[0],
                (List<String>) args[1]);
        case "snapshotSource" ->
            controller.snapshotSource((XSSFWorkbook) args[0], (XSSFPivotTable) args[1]);
        case "sourceColumnName" ->
            ExcelPivotTableSourceSupport.sourceColumnName(
                (List<String>) args[0], (Integer) args[1]);
        case "sourceColumns" ->
            ExcelPivotTableSourceSupport.sourceColumns(
                (org.apache.poi.xssf.usermodel.XSSFSheet) args[0],
                (AreaReference) args[1],
                (String) args[2]);
        case "sourceSheetName" ->
            ExcelPivotTableSourceSupport.sourceSheetName(
                (AreaReference) args[0], (Name) args[1], (String) args[2]);
        case "tableByName" ->
            ExcelPivotTableSourceSupport.tableByName(
                (XSSFWorkbook) args[0], (String) args[1], (String) args[2]);
        case "workbookPivotCache" ->
            controller.workbookPivotCache((XSSFWorkbook) args[0], (Long) args[1]);
        default -> throw new IllegalArgumentException("Unsupported helper invocation: " + name);
      };
    }
    if (target instanceof SourceColumns sourceColumns) {
      if ("relativeIndex".equals(name)) {
        return sourceColumns.relativeIndex((String) args[0]);
      }
      throw new IllegalArgumentException("Unsupported helper invocation: " + name);
    }
    if (target instanceof ColumnAxisSnapshot columnAxisSnapshot) {
      if ("columnLabels".equals(name)) {
        return columnAxisSnapshot.columnLabels();
      }
      if ("valuesAxisOnColumns".equals(name)) {
        return columnAxisSnapshot.valuesAxisOnColumns();
      }
      throw new IllegalArgumentException("Unsupported helper invocation: " + name);
    }
    throw new IllegalArgumentException(
        "Unsupported helper target: " + target.getClass().getName() + "#" + name);
  }

  private static <T extends Throwable> T assertInvocationFailure(
      Class<T> type, ThrowingRunnable runnable) {
    return assertThrows(type, runnable::run);
  }

  @FunctionalInterface
  private interface ThrowingRunnable {
    void run() throws Exception;
  }

  private static final class NullDocumentPart extends POIXMLDocumentPart {}

  private static final class SyntheticDocumentPart extends POIXMLDocumentPart {
    private SyntheticDocumentPart(PackagePart packagePart) {
      super(packagePart);
    }
  }

  private static final class ThrowingPivotTable extends XSSFPivotTable {
    private final CTPivotTableDefinition definition;

    private ThrowingPivotTable(CTPivotTableDefinition definition) {
      this.definition = definition;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
      throw new IllegalStateException("broken cache relation");
    }
  }

  private static final class NoCachePivotTable extends XSSFPivotTable {
    private final CTPivotTableDefinition definition;

    private NoCachePivotTable(CTPivotTableDefinition definition) {
      this.definition = definition;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public org.apache.poi.xssf.usermodel.XSSFPivotCache getPivotCache() {
      return null;
    }
  }

  private static final class CacheOnlyPivotTable extends XSSFPivotTable {
    private final CTPivotTableDefinition definition;
    private final XSSFPivotCache pivotCache;

    private CacheOnlyPivotTable(CTPivotTableDefinition definition, XSSFPivotCache pivotCache) {
      this.definition = definition;
      this.pivotCache = pivotCache;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public XSSFPivotCache getPivotCache() {
      return pivotCache;
    }

    @Override
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
      return null;
    }
  }

  private static final class SyntheticPivotTable extends XSSFPivotTable {
    private final CTPivotTableDefinition definition;
    private final XSSFPivotCacheDefinition cacheDefinition;

    private SyntheticPivotTable(
        CTPivotTableDefinition definition, XSSFPivotCacheDefinition cacheDefinition) {
      this.definition = definition;
      this.cacheDefinition = cacheDefinition;
    }

    @Override
    public CTPivotTableDefinition getCTPivotTableDefinition() {
      return definition;
    }

    @Override
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
      return cacheDefinition;
    }
  }

  private static void removeChild(org.apache.xmlbeans.XmlObject xmlObject, String localName) {
    try (var cursor = xmlObject.newCursor()) {
      if (cursor.toChild("http://schemas.openxmlformats.org/spreadsheetml/2006/main", localName)) {
        cursor.removeXml();
      } else {
        fail("Expected child element not found: " + localName);
      }
    }
  }
}
