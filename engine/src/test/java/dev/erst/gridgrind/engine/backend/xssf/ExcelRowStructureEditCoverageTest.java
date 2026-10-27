package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.engine.layout.StructuralEdits;
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

      StructuralEdits.shiftRows(sheet, new ExcelRowSpan(1, 5), 2);

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
              IllegalArgumentException.class, () -> StructuralEdits.insertRows(tableSheet, 1, 1));
      assertTrue(tableFailure.getMessage().contains("table 'BudgetTable'"));

      XSSFSheet autofilterSheet = workbook.createSheet("Autofilter");
      seedSheetAutofilter(autofilterSheet);
      IllegalArgumentException autofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> StructuralEdits.deleteRows(autofilterSheet, new ExcelRowSpan(1, 1)));
      assertTrue(autofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet validationSheet = workbook.createSheet("Validations");
      seedDataValidation(validationSheet);
      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> StructuralEdits.shiftRows(validationSheet, new ExcelRowSpan(1, 2), 1));
      assertTrue(validationFailure.getMessage().contains("data validation"));
    }
  }

  @Test
  void rowDeleteRejectsOverlappingRangeBackedNamedRanges() throws Exception {
    assertRowNamedRangeRejected(
        "Budget",
        "BudgetWindow",
        "Budget!$A$3:$B$4",
        new ExcelRowSpan(2, 2),
        (workbook, sheet, rows) -> runUnchecked(() -> StructuralEdits.deleteRows(sheet, rows)),
        "A3",
        "Low",
        "B4",
        "High");
  }

  @Test
  void rowShiftRejectsDestructiveRangeBackedNamedRanges() throws Exception {
    assertRowNamedRangeRejected(
        "Budget",
        "BudgetWindow",
        "Budget!$A$1:$B$2",
        new ExcelRowSpan(2, 3),
        (workbook, sheet, rows) -> runUnchecked(() -> StructuralEdits.shiftRows(sheet, rows, -2)),
        "A1",
        "Named",
        "B2",
        "Range",
        "A3",
        "Shifted",
        "A4",
        "Rows");
  }

  @Test
  void rowShiftAllowsUntouchedRangeBackedNamedRangesBetweenSourceAndDestination() throws Exception {
    assertRowNamedRangeUntouched(
        "Budget",
        "UntouchedRows",
        "Budget!$A$6:$A$7",
        new ExcelRowSpan(0, 0),
        (workbook, sheet, rows) -> runUnchecked(() -> StructuralEdits.shiftRows(sheet, rows, 10)),
        "A1",
        "Move",
        "A6",
        "Named",
        "A7",
        "Range",
        "A11",
        "Tail");
  }

  @Test
  void rowShiftRejectsPartiallyMovedRangeBackedNamedRanges() throws Exception {
    assertRowNamedRangeRejected(
        "Budget",
        "BudgetWindow",
        "Budget!$A$2:$B$3",
        new ExcelRowSpan(2, 3),
        (workbook, sheet, rows) -> runUnchecked(() -> StructuralEdits.shiftRows(sheet, rows, -2)),
        "A2",
        "Named",
        "B3",
        "Range",
        "A4",
        "Rows");
  }

  @Test
  void rowDeleteAndShiftIgnoreBlankNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      seedBlankNamedRange(workbook);

      XSSFSheet deleteSheet = workbook.createSheet("DeleteRows");
      setString(deleteSheet, "A1", "Header");
      setString(deleteSheet, "A4", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.deleteRows(deleteSheet, new ExcelRowSpan(2, 2)));

      XSSFSheet shiftSheet = workbook.createSheet("ShiftRows");
      setString(shiftSheet, "A11", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.shiftRows(shiftSheet, new ExcelRowSpan(10, 10), 1));
    }
  }

  @Test
  void deleteRowsRejectsEmptyAndOutOfBoundsSheetsAndRemovesTrailingRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet emptySheet = workbook.createSheet("EmptyRows");
      IllegalArgumentException emptyFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> StructuralEdits.deleteRows(emptySheet, new ExcelRowSpan(0, 0)));
      assertTrue(emptyFailure.getMessage().contains("requires at least one existing row"));

      XSSFSheet boundsSheet = workbook.createSheet("BoundsRows");
      setString(boundsSheet, "A1", "Header");
      IllegalArgumentException boundsFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> StructuralEdits.deleteRows(boundsSheet, new ExcelRowSpan(1, 1)));
      assertTrue(boundsFailure.getMessage().contains("last existing row is 0 (Excel row 1)"));
      assertTrue(boundsFailure.getMessage().contains("requested lastRowIndex 1 (Excel row 2)"));

      XSSFSheet deleteSheet = workbook.createSheet("DeleteRows");
      setString(deleteSheet, "A1", "Keep");
      setString(deleteSheet, "A2", "Drop");

      StructuralEdits.deleteRows(deleteSheet, new ExcelRowSpan(1, 1));

      assertEquals(0, deleteSheet.getLastRowNum());
      assertNull(deleteSheet.getRow(1));
    }
  }

  @Test
  void insertRowsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("InsertTableRows");
      seedTable(tableSheet, workbook, "InsertRowsTable");

      assertDoesNotThrow(() -> StructuralEdits.insertRows(tableSheet, 3, 1));

      XSSFSheet autofilterSheet = workbook.createSheet("InsertAutofilterRowsSafe");
      seedSheetAutofilter(autofilterSheet);

      assertDoesNotThrow(() -> StructuralEdits.insertRows(autofilterSheet, 2, 1));

      XSSFSheet validationSheet = workbook.createSheet("InsertValidationRowsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "A5", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.insertRows(validationSheet, 4, 1));
    }
  }

  @Test
  void deleteRowsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("DeleteTableRowsSafe");
      seedTable(tableSheet, workbook, "DeleteRowsTable");
      setString(tableSheet, "A4", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.deleteRows(tableSheet, new ExcelRowSpan(3, 3)));

      XSSFSheet autofilterSheet = workbook.createSheet("DeleteAutofilterRowsSafe");
      seedSheetAutofilter(autofilterSheet);
      setString(autofilterSheet, "A3", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.deleteRows(autofilterSheet, new ExcelRowSpan(2, 2)));

      XSSFSheet validationSheet = workbook.createSheet("DeleteValidationRowsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "A5", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.deleteRows(validationSheet, new ExcelRowSpan(4, 4)));
    }
  }

  @Test
  void shiftRowsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("ShiftTableRowsSafe");
      seedTable(tableSheet, workbook, "ShiftRowsTable");
      setString(tableSheet, "A11", "Tail");

      assertDoesNotThrow(() -> StructuralEdits.shiftRows(tableSheet, new ExcelRowSpan(10, 10), 1));

      XSSFSheet autofilterSheet = workbook.createSheet("ShiftAutofilterRowsSafe");
      seedSheetAutofilter(autofilterSheet);
      setString(autofilterSheet, "A11", "Tail");

      assertDoesNotThrow(
          () -> StructuralEdits.shiftRows(autofilterSheet, new ExcelRowSpan(10, 10), 1));

      XSSFSheet validationSheet = workbook.createSheet("ShiftValidationRowsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "A11", "Tail");

      assertDoesNotThrow(
          () -> StructuralEdits.shiftRows(validationSheet, new ExcelRowSpan(10, 10), 1));
    }
  }

  @Test
  void insertRowsAndColumnsAllowAppendingAtExistingEdge() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet appendRowsSheet = workbook.createSheet("AppendRows");
      setString(appendRowsSheet, "A1", "Header");

      assertDoesNotThrow(() -> StructuralEdits.insertRows(appendRowsSheet, 1, 1));
      assertEquals(0, appendRowsSheet.getLastRowNum());

      XSSFSheet appendColumnsSheet = workbook.createSheet("AppendColumns");
      setString(appendColumnsSheet, "A1", "Header");

      assertDoesNotThrow(() -> StructuralEdits.insertColumns(appendColumnsSheet, 1, 1));
      assertEquals(0, StructuralEdits.lastColumnIndex(appendColumnsSheet));
    }
  }

  @Test
  void insertingRowsAndColumnsMovesPictureAnchorsWithTheirCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Drawing");
      sheet.createRow(6).createCell(5).setCellValue("tail");
      var anchor = new org.apache.poi.xssf.usermodel.XSSFClientAnchor(0, 0, 0, 0, 1, 2, 3, 4);
      anchor.setAnchorType(org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.MOVE_AND_RESIZE);
      byte[] png =
          java.util.Base64.getDecoder()
              .decode(
                  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(
          anchor, workbook.addPicture(png, org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PNG));
      StructuralEdits.insertRows(sheet, 2, 2);
      StructuralEdits.insertColumns(sheet, 1, 2);
      var bytes = new java.io.ByteArrayOutputStream();
      workbook.write(bytes);
      try (var reopened = new XSSFWorkbook(new java.io.ByteArrayInputStream(bytes.toByteArray()))) {
        var picture = reopened.getSheetAt(0).getDrawingPatriarch().getShapes().getFirst();
        var persisted = (org.apache.poi.xssf.usermodel.XSSFClientAnchor) picture.getAnchor();
        assertEquals(4, persisted.getRow1());
        assertEquals(6, persisted.getRow2());
        assertEquals(3, persisted.getCol1());
        assertEquals(5, persisted.getCol2());
      }
    }
  }
}
