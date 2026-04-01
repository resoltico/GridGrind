package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for direct sheet-autofilter authoring, introspection, and health analysis. */
class ExcelAutofilterControllerTest {
  private final ExcelAutofilterController controller = new ExcelAutofilterController();

  @Test
  void setSheetAutofilter_roundTripsOwnedMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
          controller.sheetOwnedAutofilters(sheet));
      assertEquals(1, controller.sheetAutofilterCount(sheet));

      ByteArrayOutputStream output = new ByteArrayOutputStream();
      workbook.write(output);
      try (XSSFWorkbook reopened =
          new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
        XSSFSheet reopenedSheet = reopened.getSheet("Ops");
        assertEquals(
            List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
            controller.sheetOwnedAutofilters(reopenedSheet));
      }
    }
  }

  @Test
  void clearSheetAutofilter_removesFilterDatabaseNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      controller.setSheetAutofilter(sheet, "A1:B3");

      assertTrue(
          sheet.getWorkbook().getAllNames().stream()
              .anyMatch(ExcelAutofilterControllerTest::isFilterDatabaseName));

      controller.clearSheetAutofilter(sheet);

      assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
      assertTrue(
          sheet.getWorkbook().getAllNames().stream()
              .noneMatch(ExcelAutofilterControllerTest::isFilterDatabaseName));
    }
  }

  @Test
  void clearSheetAutofilter_removesOnlySameSheetFilterDatabaseNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet ops = populatedSheet(workbook);
      XSSFSheet archive = populatedSheet(workbook, "Archive");
      controller.setSheetAutofilter(ops, "A1:B3");

      Name archiveFilter = workbook.createName();
      archiveFilter.setNameName("_xlnm._FilterDatabase");
      archiveFilter.setSheetIndex(workbook.getSheetIndex(archive));
      archiveFilter.setRefersToFormula("Archive!$A$1:$B$3");

      Name unrelated = workbook.createName();
      unrelated.setNameName("OpsRange");
      unrelated.setSheetIndex(workbook.getSheetIndex(ops));
      unrelated.setRefersToFormula("Ops!$A$1:$A$3");

      controller.clearSheetAutofilter(ops);

      assertTrue(workbook.getAllNames().contains(archiveFilter));
      assertTrue(workbook.getAllNames().contains(unrelated));
      assertFalse(
          workbook.getAllNames().stream()
              .anyMatch(
                  name ->
                      name.getSheetIndex() == workbook.getSheetIndex(ops)
                          && isFilterDatabaseName(name)));
    }
  }

  @Test
  void absentSheetAutofilterReadsAsEmptyAndClearsAsNoOp() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);

      assertEquals(List.of(), controller.sheetOwnedAutofilters(sheet));
      assertEquals(0, controller.sheetAutofilterCount(sheet));

      controller.clearSheetAutofilter(sheet);

      assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
      assertEquals(List.of(), controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void sheetOwnedAutofilters_readsBlankRefAsEmptyString() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getCTWorksheet().addNewAutoFilter().setRef("");

      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("")),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void setSheetAutofilter_rejectsBlankHeaderRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.setSheetAutofilter(sheet, "A1:B2"));

      assertEquals(
          "autofilter range must include a nonblank header row: A1:B2", failure.getMessage());
    }
  }

  @Test
  void setSheetAutofilter_rejectsOverlapWithExistingTables() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.setSheetAutofilter(sheet, "A1:B3"));

      assertEquals(
          "sheet-level autofilter range must not overlap an existing table range: Table1@A1:B3",
          failure.getMessage());
    }
  }

  @Test
  void sheetAutofilterHealthFlagsInvalidRangesAndTableOverlap() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getCTWorksheet().addNewAutoFilter().setRef("A0:B3");

      List<WorkbookAnalysis.AnalysisFinding> invalidRangeFindings =
          controller.sheetAutofilterHealthFindings("Ops", sheet, List.of());

      assertEquals(1, invalidRangeFindings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
          invalidRangeFindings.getFirst().code());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));
      sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));

      List<WorkbookAnalysis.AnalysisFinding> overlapFindings =
          controller.sheetAutofilterHealthFindings(
              "Ops",
              sheet,
              List.of(
                  new ExcelTableSnapshot(
                      "Table1",
                      "Ops",
                      "A1:B3",
                      1,
                      0,
                      List.of("Owner", "Task"),
                      new ExcelTableStyleSnapshot.None(),
                      false)));

      assertEquals(1, overlapFindings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          overlapFindings.getFirst().code());
    }
  }

  @Test
  void sheetAutofilterHealthFlagsBlankHeaderRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getRow(0).getCell(0).setBlank();
      sheet.getRow(0).getCell(1).setBlank();
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.sheetAutofilterHealthFindings("Ops", sheet, List.of());

      assertEquals(1, findings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          findings.getFirst().code());
    }
  }

  @Test
  void sheetAutofilterHealthIgnoresInvalidAndNonOverlappingTables() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.sheetAutofilterHealthFindings(
              "Ops",
              sheet,
              List.of(
                  new ExcelTableSnapshot(
                      "Invalid",
                      "Ops",
                      "A0:B3",
                      1,
                      0,
                      List.of("Owner", "Task"),
                      new ExcelTableStyleSnapshot.None(),
                      false),
                  new ExcelTableSnapshot(
                      "Desk",
                      "Ops",
                      "D1:E3",
                      1,
                      0,
                      List.of("Desk", "Region"),
                      new ExcelTableStyleSnapshot.None(),
                      false)));

      assertEquals(List.of(), findings);
    }
  }

  @Test
  void setSheetAutofilter_allowsInvalidExistingTableMetadataWhenItDoesNotOverlap()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      XSSFTable table = sheet.createTable(new AreaReference("D1:E3", SpreadsheetVersion.EXCEL2007));
      table.getCTTable().setRef("D0:E3");

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void setSheetAutofilter_allowsValidNonOverlappingTableRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getRow(0).createCell(3).setCellValue("Desk");
      sheet.getRow(0).createCell(4).setCellValue("Region");
      sheet.getRow(1).createCell(3).setCellValue("A1");
      sheet.getRow(1).createCell(4).setCellValue("North");
      sheet.getRow(2).createCell(3).setCellValue("B1");
      sheet.getRow(2).createCell(4).setCellValue("South");
      sheet.createTable(new AreaReference("D1:E3", SpreadsheetVersion.EXCEL2007));

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  private static XSSFSheet populatedSheet(XSSFWorkbook workbook) {
    return populatedSheet(workbook, "Ops");
  }

  private static XSSFSheet populatedSheet(XSSFWorkbook workbook, String name) {
    XSSFSheet sheet = workbook.createSheet(name);
    sheet.createRow(0).createCell(0).setCellValue("Owner");
    sheet.getRow(0).createCell(1).setCellValue("Task");
    sheet.createRow(1).createCell(0).setCellValue("Ada");
    sheet.getRow(1).createCell(1).setCellValue("Queue");
    sheet.createRow(2).createCell(0).setCellValue("Lin");
    sheet.getRow(2).createCell(1).setCellValue("Pack");
    return sheet;
  }

  private static boolean isFilterDatabaseName(org.apache.poi.ss.usermodel.Name name) {
    return "_XLNM._FILTERDATABASE".equalsIgnoreCase(name.getNameName());
  }
}
