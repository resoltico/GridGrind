package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Row structural edit and guard coverage. */
class ExcelRowStructureEditCoverageTest extends ExcelRowColumnStructureTestSupport {
  @Test
  void shiftRowsPreservesSupportedStructures() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      seedSupportedScenario(workbook, sheet, true);

      controller.shiftRows(sheet, new ExcelRowSpan(1, 5), 2);

      assertEquals("Budget!$B$4:$B$6", workbook.getName("BudgetValues").getRefersToFormula());
      assertEquals("E4:F5", sheet.getMergedRegion(0).formatAsString());
      assertEquals("SUM(B4:B6)", sheet.getRow(3).getCell(6).getCellFormula());
      assertEquals("SUM(B4:B6)", sheet.getRow(11).getCell(7).getCellFormula());
      assertEquals("D7", hyperlinkAddresses(sheet).getFirst());
      assertEquals("D7", commentAddresses(sheet).getFirst());
      assertEquals(
          "B4:B7",
          sheet
              .getSheetConditionalFormatting()
              .getConditionalFormattingAt(0)
              .getFormattingRanges()[0]
              .formatAsString());
    }
  }

  @Test
  void rowStructuralEditsRejectAffectedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("Tables");
      seedTable(tableSheet, workbook);
      IllegalArgumentException tableFailure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.insertRows(tableSheet, 1, 1));
      assertTrue(tableFailure.getMessage().contains("table 'BudgetTable'"));

      XSSFSheet autofilterSheet = workbook.createSheet("Autofilter");
      seedSheetAutofilter(autofilterSheet);
      IllegalArgumentException autofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteRows(autofilterSheet, new ExcelRowSpan(1, 1)));
      assertTrue(autofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet validationSheet = workbook.createSheet("Validations");
      seedDataValidation(validationSheet);
      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftRows(validationSheet, new ExcelRowSpan(1, 2), 1));
      assertTrue(validationFailure.getMessage().contains("data validation"));
    }
  }

  @Test
  void rowDeleteRejectsOverlappingRangeBackedNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A3", "Low");
      setString(sheet, "B4", "High");
      seedNamedRange(workbook, "BudgetWindow", "Budget!$A$3:$B$4");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteRows(sheet, new ExcelRowSpan(2, 2)));
      assertTrue(failure.getMessage().contains("named range 'BudgetWindow'"));
    }
  }

  @Test
  void rowShiftRejectsDestructiveRangeBackedNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Named");
      setString(sheet, "B2", "Range");
      setString(sheet, "A3", "Shifted");
      setString(sheet, "A4", "Rows");
      seedNamedRange(workbook, "BudgetWindow", "Budget!$A$1:$B$2");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftRows(sheet, new ExcelRowSpan(2, 3), -2));
      assertTrue(failure.getMessage().contains("named range 'BudgetWindow'"));
    }
  }

  @Test
  void rowShiftAllowsUntouchedRangeBackedNamedRangesBetweenSourceAndDestination() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Move");
      setString(sheet, "A6", "Named");
      setString(sheet, "A7", "Range");
      setString(sheet, "A11", "Tail");
      seedNamedRange(workbook, "UntouchedRows", "Budget!$A$6:$A$7");

      assertDoesNotThrow(() -> controller.shiftRows(sheet, new ExcelRowSpan(0, 0), 10));
      assertEquals("Budget!$A$6:$A$7", workbook.getName("UntouchedRows").getRefersToFormula());
    }
  }

  @Test
  void rowShiftRejectsPartiallyMovedRangeBackedNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A2", "Named");
      setString(sheet, "B3", "Range");
      setString(sheet, "A4", "Rows");
      seedNamedRange(workbook, "BudgetWindow", "Budget!$A$2:$B$3");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftRows(sheet, new ExcelRowSpan(2, 3), -2));
      assertTrue(failure.getMessage().contains("named range 'BudgetWindow'"));
    }
  }

  @Test
  void rowDeleteAndShiftIgnoreBlankNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      seedBlankNamedRange(workbook);

      XSSFSheet deleteSheet = workbook.createSheet("DeleteRows");
      setString(deleteSheet, "A1", "Header");
      setString(deleteSheet, "A4", "Tail");

      assertDoesNotThrow(() -> controller.deleteRows(deleteSheet, new ExcelRowSpan(2, 2)));

      XSSFSheet shiftSheet = workbook.createSheet("ShiftRows");
      setString(shiftSheet, "A11", "Tail");

      assertDoesNotThrow(() -> controller.shiftRows(shiftSheet, new ExcelRowSpan(10, 10), 1));
    }
  }

  @Test
  void deleteRowsRejectsEmptyAndOutOfBoundsSheetsAndRemovesTrailingRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet emptySheet = workbook.createSheet("EmptyRows");
      IllegalArgumentException emptyFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteRows(emptySheet, new ExcelRowSpan(0, 0)));
      assertTrue(emptyFailure.getMessage().contains("requires at least one existing row"));

      XSSFSheet boundsSheet = workbook.createSheet("BoundsRows");
      setString(boundsSheet, "A1", "Header");
      IllegalArgumentException boundsFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteRows(boundsSheet, new ExcelRowSpan(1, 1)));
      assertTrue(boundsFailure.getMessage().contains("last existing row is 0 (Excel row 1)"));
      assertTrue(boundsFailure.getMessage().contains("requested lastRowIndex 1 (Excel row 2)"));

      XSSFSheet deleteSheet = workbook.createSheet("DeleteRows");
      setString(deleteSheet, "A1", "Keep");
      setString(deleteSheet, "A2", "Drop");

      controller.deleteRows(deleteSheet, new ExcelRowSpan(1, 1));

      assertEquals(0, deleteSheet.getLastRowNum());
      assertNull(deleteSheet.getRow(1));
    }
  }

  @Test
  void insertRowsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("InsertTableRows");
      seedTable(tableSheet, workbook, "InsertRowsTable");

      assertDoesNotThrow(() -> controller.insertRows(tableSheet, 3, 1));

      XSSFSheet autofilterSheet = workbook.createSheet("InsertAutofilterRowsSafe");
      seedSheetAutofilter(autofilterSheet);

      assertDoesNotThrow(() -> controller.insertRows(autofilterSheet, 2, 1));

      XSSFSheet validationSheet = workbook.createSheet("InsertValidationRowsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "A5", "Tail");

      assertDoesNotThrow(() -> controller.insertRows(validationSheet, 4, 1));
    }
  }

  @Test
  void deleteRowsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("DeleteTableRowsSafe");
      seedTable(tableSheet, workbook, "DeleteRowsTable");
      setString(tableSheet, "A4", "Tail");

      assertDoesNotThrow(() -> controller.deleteRows(tableSheet, new ExcelRowSpan(3, 3)));

      XSSFSheet autofilterSheet = workbook.createSheet("DeleteAutofilterRowsSafe");
      seedSheetAutofilter(autofilterSheet);
      setString(autofilterSheet, "A3", "Tail");

      assertDoesNotThrow(() -> controller.deleteRows(autofilterSheet, new ExcelRowSpan(2, 2)));

      XSSFSheet validationSheet = workbook.createSheet("DeleteValidationRowsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "A5", "Tail");

      assertDoesNotThrow(() -> controller.deleteRows(validationSheet, new ExcelRowSpan(4, 4)));
    }
  }

  @Test
  void shiftRowsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("ShiftTableRowsSafe");
      seedTable(tableSheet, workbook, "ShiftRowsTable");
      setString(tableSheet, "A11", "Tail");

      assertDoesNotThrow(() -> controller.shiftRows(tableSheet, new ExcelRowSpan(10, 10), 1));

      XSSFSheet autofilterSheet = workbook.createSheet("ShiftAutofilterRowsSafe");
      seedSheetAutofilter(autofilterSheet);
      setString(autofilterSheet, "A11", "Tail");

      assertDoesNotThrow(() -> controller.shiftRows(autofilterSheet, new ExcelRowSpan(10, 10), 1));

      XSSFSheet validationSheet = workbook.createSheet("ShiftValidationRowsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "A11", "Tail");

      assertDoesNotThrow(() -> controller.shiftRows(validationSheet, new ExcelRowSpan(10, 10), 1));
    }
  }

  @Test
  void insertRowsAndColumnsAllowAppendingAtExistingEdge() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet appendRowsSheet = workbook.createSheet("AppendRows");
      setString(appendRowsSheet, "A1", "Header");

      assertDoesNotThrow(() -> controller.insertRows(appendRowsSheet, 1, 1));
      assertEquals(0, appendRowsSheet.getLastRowNum());

      XSSFSheet appendColumnsSheet = workbook.createSheet("AppendColumns");
      setString(appendColumnsSheet, "A1", "Header");

      assertDoesNotThrow(() -> controller.insertColumns(appendColumnsSheet, 1, 1));
      assertEquals(0, controller.lastColumnIndex(appendColumnsSheet));
    }
  }
}
