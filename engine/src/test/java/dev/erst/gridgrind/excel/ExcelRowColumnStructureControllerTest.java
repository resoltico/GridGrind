package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;

/** Tests for structural row and column editing plus layout normalization. */
@SuppressWarnings("StringConcatToTextBlock")
class ExcelRowColumnStructureControllerTest {
  private final ExcelRowColumnStructureController controller =
      new ExcelRowColumnStructureController();

  @Test
  void layoutReportsHiddenOutlineAndCollapsedFacts() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), true);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), true);
      controller.setRowVisibility(sheet, new ExcelRowSpan(5, 5), true);
      controller.setColumnVisibility(sheet, new ExcelColumnSpan(5, 5), true);

      List<WorkbookReadResult.RowLayout> rows = controller.rowLayouts(sheet);
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);

      assertEquals(6, rows.size());
      assertTrue(rows.get(1).hidden());
      assertEquals(1, rows.get(1).outlineLevel());
      assertTrue(rows.get(4).collapsed());
      assertTrue(rows.get(5).hidden());

      assertEquals(6, columns.size());
      assertTrue(columns.get(1).hidden());
      assertEquals(1, columns.get(1).outlineLevel());
      assertTrue(columns.get(4).collapsed());
      assertTrue(columns.get(5).hidden());
    }
  }

  @Test
  void collapsedRowGroupsPersistControlRowsAtSparseSheetTail() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-row-group-tail-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.groupRows(sheet, new ExcelRowSpan(0, 1), true);

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookReadResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(3, layout.rows().size());
    assertTrue(layout.rows().get(0).hidden());
    assertEquals(1, layout.rows().get(0).outlineLevel());
    assertTrue(layout.rows().get(1).hidden());
    assertEquals(1, layout.rows().get(1).outlineLevel());
    assertFalse(layout.rows().get(2).hidden());
    assertTrue(layout.rows().get(2).collapsed());
  }

  @Test
  void redundantNoOpColumnUngroupsDoNotMaterializeGhostColumnMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));

      assertTrue(controller.columnLayouts(sheet).isEmpty());
      assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());
      assertEquals(0, sheet.getCTWorksheet().getColsArray(0).sizeOfColArray());
    }
  }

  @Test
  void partialColumnUngroupPreservesRemainingCollapsedBandAcrossRoundTrip() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-column-ungroup-collapse-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.groupColumns(sheet, new ExcelColumnSpan(0, 3), true);
      controller.ungroupColumns(sheet, new ExcelColumnSpan(1, 1));

      List<WorkbookReadResult.ColumnLayout> inMemoryColumns = controller.columnLayouts(sheet);
      assertEquals(5, inMemoryColumns.size(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(0).hidden(), "in-memory columns=" + inMemoryColumns);
      assertFalse(inMemoryColumns.get(1).collapsed(), "in-memory columns=" + inMemoryColumns);
      assertEquals(
          0, inMemoryColumns.get(1).outlineLevel(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(2).hidden(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(3).hidden(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(4).collapsed(), "in-memory columns=" + inMemoryColumns);

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookReadResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(5, layout.columns().size(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(0).hidden(), "reopened columns=" + layout.columns());
    assertFalse(layout.columns().get(1).collapsed(), "reopened columns=" + layout.columns());
    assertEquals(0, layout.columns().get(1).outlineLevel(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(2).hidden(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(3).hidden(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(4).collapsed(), "reopened columns=" + layout.columns());
  }

  @Test
  void ungroupedSparseRowsNormalizeOutlineLevelToZeroAfterReopen() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-row-ungroup-tail-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.ungroupRows(sheet, new ExcelRowSpan(1, 3));

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookReadResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(4, layout.rows().size());
    assertEquals(0, layout.rows().get(1).outlineLevel());
    assertEquals(0, layout.rows().get(2).outlineLevel());
    assertEquals(0, layout.rows().get(3).outlineLevel());
  }

  @Test
  void groupRowsAndColumnsWithoutCollapseKeepsBandsExpanded() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), false);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), false);

      List<WorkbookReadResult.RowLayout> rows = controller.rowLayouts(sheet);
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);

      assertEquals(4, rows.size());
      assertFalse(rows.get(1).hidden());
      assertEquals(1, rows.get(1).outlineLevel());
      assertFalse(rows.get(3).collapsed());

      assertEquals(4, columns.size());
      assertFalse(columns.get(1).hidden());
      assertEquals(1, columns.get(1).outlineLevel());
      assertFalse(columns.get(3).collapsed());
    }
  }

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
  void deleteColumnsRejectsEmptyAndOutOfBoundsSheets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet emptySheet = workbook.createSheet("EmptyColumns");
      IllegalArgumentException emptyFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteColumns(emptySheet, new ExcelColumnSpan(0, 0)));
      assertTrue(emptyFailure.getMessage().contains("requires at least one existing column"));

      XSSFSheet boundsSheet = workbook.createSheet("BoundsColumns");
      setString(boundsSheet, "A1", "Header");
      IllegalArgumentException boundsFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteColumns(boundsSheet, new ExcelColumnSpan(1, 1)));
      assertTrue(boundsFailure.getMessage().contains("last existing column is 0 (Excel column A)"));
      assertTrue(
          boundsFailure.getMessage().contains("requested lastColumnIndex 1 (Excel column B)"));
    }
  }

  @Test
  void insertAndShiftEditsRejectOutOfBoundsIndices() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet insertRowBounds = workbook.createSheet("InsertRowBounds");
      IllegalArgumentException rowGapFailure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.insertRows(insertRowBounds, 1, 1));
      assertTrue(rowGapFailure.getMessage().contains("rowIndex 1 (Excel row 2)"));
      assertTrue(rowGapFailure.getMessage().contains("last existing row + 1: 0 (Excel row 1)"));

      XSSFSheet insertRowLimit = workbook.createSheet("InsertRowLimit");
      insertRowLimit.createRow(ExcelRowSpan.MAX_ROW_INDEX);
      IllegalArgumentException rowOverflowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.insertRows(insertRowLimit, ExcelRowSpan.MAX_ROW_INDEX, 2));
      assertTrue(
          rowOverflowFailure
              .getMessage()
              .contains("destination last row would be 1048576 (Excel row 1048577)"));
      assertTrue(
          rowOverflowFailure.getMessage().contains("maximum is 1048575 (Excel row 1048576)"));

      XSSFSheet insertColumnBounds = workbook.createSheet("InsertColumnBounds");
      setString(insertColumnBounds, "A1", "Header");
      IllegalArgumentException columnGapFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.insertColumns(insertColumnBounds, 2, 1));
      assertTrue(columnGapFailure.getMessage().contains("columnIndex 2 (Excel column C)"));
      assertTrue(
          columnGapFailure.getMessage().contains("last existing column + 1: 1 (Excel column B)"));

      XSSFSheet insertColumnLimit = workbook.createSheet("InsertColumnLimit");
      setString(insertColumnLimit, "XFD1", "Edge");
      IllegalArgumentException columnOverflowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.insertColumns(insertColumnLimit, ExcelColumnSpan.MAX_COLUMN_INDEX, 2));
      assertTrue(
          columnOverflowFailure
              .getMessage()
              .contains("destination last column would be 16384 (Excel column XFE)"));
      assertTrue(
          columnOverflowFailure.getMessage().contains("maximum is 16383 (Excel column XFD)"));

      IllegalArgumentException shiftRowLowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftRows(insertRowBounds, new ExcelRowSpan(0, 0), -1));
      assertTrue(
          shiftRowLowFailure.getMessage().contains("firstRowIndex 0 (Excel row 1) by delta -1"));
      assertTrue(
          shiftRowLowFailure.getMessage().contains("before the first worksheet row (Excel row 1)"));

      IllegalArgumentException shiftRowHighFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.shiftRows(
                      insertRowBounds,
                      new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX, ExcelRowSpan.MAX_ROW_INDEX),
                      1));
      assertTrue(
          shiftRowHighFailure
              .getMessage()
              .contains("lastRowIndex 1048575 (Excel row 1048576) by delta 1"));
      assertTrue(
          shiftRowHighFailure.getMessage().contains("maximum row 1048575 (Excel row 1048576)"));

      IllegalArgumentException shiftColumnLowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftColumns(insertColumnBounds, new ExcelColumnSpan(0, 0), -1));
      assertTrue(
          shiftColumnLowFailure
              .getMessage()
              .contains("firstColumnIndex 0 (Excel column A) by delta -1"));
      assertTrue(
          shiftColumnLowFailure
              .getMessage()
              .contains("before the first worksheet column (Excel column A)"));

      IllegalArgumentException shiftColumnHighFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.shiftColumns(
                      insertColumnBounds,
                      new ExcelColumnSpan(
                          ExcelColumnSpan.MAX_COLUMN_INDEX, ExcelColumnSpan.MAX_COLUMN_INDEX),
                      1));
      assertTrue(
          shiftColumnHighFailure
              .getMessage()
              .contains("lastColumnIndex 16383 (Excel column XFD) by delta 1"));
      assertTrue(
          shiftColumnHighFailure.getMessage().contains("maximum column 16383 (Excel column XFD)"));
    }
  }

  @Test
  void shiftColumnsPreservesSupportedStructuresOnFormulaFreeWorkbook() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      seedSupportedScenario(workbook, sheet, false);

      controller.shiftColumns(sheet, new ExcelColumnSpan(0, 5), 2);

      assertEquals("Budget!$D$2:$D$4", workbook.getName("BudgetValues").getRefersToFormula());
      assertEquals("G2:H3", sheet.getMergedRegion(0).formatAsString());
      assertEquals("F5", hyperlinkAddresses(sheet).getFirst());
      assertEquals("F5", commentAddresses(sheet).getFirst());
      assertEquals(
          "D2:D5",
          sheet
              .getSheetConditionalFormatting()
              .getConditionalFormattingAt(0)
              .getFormattingRanges()[0]
              .formatAsString());
      assertEquals("Hosting", sheet.getRow(1).getCell(2).getStringCellValue());
      assertEquals(42.0, sheet.getRow(1).getCell(3).getNumericCellValue());
    }
  }

  @Test
  void columnStructuralEditsRejectFormulasAndFormulaDefinedNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet formulaSheet = workbook.createSheet("FormulaSheet");
      seedSupportedScenario(workbook, formulaSheet, true);
      IllegalArgumentException formulaFailure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.insertColumns(formulaSheet, 1, 1));
      assertTrue(formulaFailure.getMessage().contains("formulas are present"));

      XSSFWorkbook namedWorkbook = new XSSFWorkbook();
      try (namedWorkbook) {
        XSSFSheet namedSheet = namedWorkbook.createSheet("NamedSheet");
        setString(namedSheet, "A1", "Header");
        Name name = namedWorkbook.createName();
        name.setNameName("DynamicBudget");
        name.setRefersToFormula("OFFSET(NamedSheet!$A$1,0,0,2,1)");

        IllegalArgumentException nameFailure =
            assertThrows(
                IllegalArgumentException.class, () -> controller.insertColumns(namedSheet, 0, 1));
        assertTrue(nameFailure.getMessage().contains("formula-defined names are present"));
      }
    }
  }

  @Test
  void columnDeleteRejectsOverlappingRangeBackedNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "C1", "Low");
      setString(sheet, "D2", "High");
      seedNamedRange(workbook, "BudgetWindow", "Budget!$C$1:$D$2");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteColumns(sheet, new ExcelColumnSpan(2, 2)));
      assertTrue(failure.getMessage().contains("named range 'BudgetWindow'"));
    }
  }

  @Test
  void columnShiftRejectsDestructiveRangeBackedNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Named");
      setString(sheet, "B2", "Range");
      setString(sheet, "C1", "Shifted");
      setString(sheet, "D1", "Columns");
      seedNamedRange(workbook, "BudgetWindow", "Budget!$A$1:$B$2");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftColumns(sheet, new ExcelColumnSpan(2, 3), -2));
      assertTrue(failure.getMessage().contains("named range 'BudgetWindow'"));
    }
  }

  @Test
  void columnShiftAllowsUntouchedRangeBackedNamedRangesBetweenSourceAndDestination()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Move");
      setString(sheet, "F1", "Named");
      setString(sheet, "G1", "Range");
      setString(sheet, "K1", "Tail");
      seedNamedRange(workbook, "UntouchedColumns", "Budget!$F$1:$G$1");

      assertDoesNotThrow(() -> controller.shiftColumns(sheet, new ExcelColumnSpan(0, 0), 10));
      assertEquals("Budget!$F$1:$G$1", workbook.getName("UntouchedColumns").getRefersToFormula());
    }
  }

  @Test
  void columnShiftRejectsPartiallyMovedRangeBackedNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "B1", "Named");
      setString(sheet, "C2", "Range");
      setString(sheet, "D1", "Columns");
      seedNamedRange(workbook, "BudgetWindow", "Budget!$B$1:$C$2");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftColumns(sheet, new ExcelColumnSpan(2, 3), -2));
      assertTrue(failure.getMessage().contains("named range 'BudgetWindow'"));
    }
  }

  @Test
  void columnDeleteAndShiftIgnoreBlankNamedRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      seedBlankNamedRange(workbook);

      XSSFSheet deleteSheet = workbook.createSheet("DeleteColumns");
      setString(deleteSheet, "A1", "Header");
      setString(deleteSheet, "D1", "Tail");

      assertDoesNotThrow(() -> controller.deleteColumns(deleteSheet, new ExcelColumnSpan(2, 2)));

      XSSFSheet shiftSheet = workbook.createSheet("ShiftColumns");
      setString(shiftSheet, "K1", "Tail");

      assertDoesNotThrow(() -> controller.shiftColumns(shiftSheet, new ExcelColumnSpan(10, 10), 1));
    }
  }

  @Test
  void columnStructuralEditsRejectAffectedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("Tables");
      seedTable(tableSheet, workbook);
      IllegalArgumentException tableFailure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.insertColumns(tableSheet, 1, 1));
      assertTrue(tableFailure.getMessage().contains("table 'BudgetTable'"));

      XSSFSheet autofilterSheet = workbook.createSheet("Autofilter");
      seedSheetAutofilter(autofilterSheet);
      IllegalArgumentException autofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteColumns(autofilterSheet, new ExcelColumnSpan(1, 1)));
      assertTrue(autofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet validationSheet = workbook.createSheet("Validations");
      seedDataValidation(validationSheet);
      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftColumns(validationSheet, new ExcelColumnSpan(0, 0), 1));
      assertTrue(validationFailure.getMessage().contains("data validation"));
    }
  }

  @Test
  void rowStructuralEditsCoverRemainingAutofilterTableAndValidationRejections() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet insertAutofilterSheet = workbook.createSheet("InsertAutofilterRows");
      seedSheetAutofilter(insertAutofilterSheet);
      IllegalArgumentException insertAutofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.insertRows(insertAutofilterSheet, 1, 1));
      assertTrue(insertAutofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet insertValidationSheet = workbook.createSheet("InsertValidationRows");
      seedDataValidation(insertValidationSheet);
      IllegalArgumentException insertValidationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.insertRows(insertValidationSheet, 1, 1));
      assertTrue(insertValidationFailure.getMessage().contains("data validation"));

      XSSFSheet deleteValidationSheet = workbook.createSheet("DeleteValidationRows");
      seedDataValidation(deleteValidationSheet);
      IllegalArgumentException deleteValidationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteRows(deleteValidationSheet, new ExcelRowSpan(1, 1)));
      assertTrue(deleteValidationFailure.getMessage().contains("data validation"));

      XSSFSheet shiftAutofilterSheet = workbook.createSheet("ShiftAutofilterRows");
      seedSheetAutofilter(shiftAutofilterSheet);
      IllegalArgumentException shiftAutofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftRows(shiftAutofilterSheet, new ExcelRowSpan(1, 1), 1));
      assertTrue(shiftAutofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet untouchedValidationSheet = workbook.createSheet("UntouchedValidationRows");
      seedDataValidation(untouchedValidationSheet);
      setString(untouchedValidationSheet, "A11", "Tail");

      assertDoesNotThrow(
          () -> controller.shiftRows(untouchedValidationSheet, new ExcelRowSpan(10, 10), 1));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet deleteTableSheet = workbook.createSheet("DeleteTableRows");
      seedTable(deleteTableSheet, workbook);
      IllegalArgumentException deleteTableFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteRows(deleteTableSheet, new ExcelRowSpan(1, 1)));
      assertTrue(deleteTableFailure.getMessage().contains("table 'BudgetTable'"));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet shiftTableSheet = workbook.createSheet("ShiftTableRows");
      seedTable(shiftTableSheet, workbook);
      IllegalArgumentException shiftTableFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftRows(shiftTableSheet, new ExcelRowSpan(1, 1), 1));
      assertTrue(shiftTableFailure.getMessage().contains("table 'BudgetTable'"));
    }
  }

  @Test
  void columnStructuralEditsCoverRemainingAutofilterTableAndValidationRejections()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet insertAutofilterSheet = workbook.createSheet("InsertAutofilterColumns");
      seedSheetAutofilter(insertAutofilterSheet);
      IllegalArgumentException insertAutofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.insertColumns(insertAutofilterSheet, 1, 1));
      assertTrue(insertAutofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet insertValidationSheet = workbook.createSheet("InsertValidationColumns");
      seedDataValidation(insertValidationSheet);
      IllegalArgumentException insertValidationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.insertColumns(insertValidationSheet, 0, 1));
      assertTrue(insertValidationFailure.getMessage().contains("data validation"));

      XSSFSheet deleteValidationSheet = workbook.createSheet("DeleteValidationColumns");
      seedDataValidation(deleteValidationSheet);
      IllegalArgumentException deleteValidationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteColumns(deleteValidationSheet, new ExcelColumnSpan(0, 0)));
      assertTrue(deleteValidationFailure.getMessage().contains("data validation"));

      XSSFSheet shiftAutofilterSheet = workbook.createSheet("ShiftAutofilterColumns");
      seedSheetAutofilter(shiftAutofilterSheet);
      IllegalArgumentException shiftAutofilterFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftColumns(shiftAutofilterSheet, new ExcelColumnSpan(0, 0), 1));
      assertTrue(shiftAutofilterFailure.getMessage().contains("sheet autofilter"));

      XSSFSheet untouchedValidationSheet = workbook.createSheet("UntouchedValidationColumns");
      seedDataValidation(untouchedValidationSheet);
      setString(untouchedValidationSheet, "K1", "Tail");

      assertDoesNotThrow(
          () -> controller.shiftColumns(untouchedValidationSheet, new ExcelColumnSpan(10, 10), 1));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet deleteTableSheet = workbook.createSheet("DeleteTableColumns");
      seedTable(deleteTableSheet, workbook);
      IllegalArgumentException deleteTableFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteColumns(deleteTableSheet, new ExcelColumnSpan(1, 1)));
      assertTrue(deleteTableFailure.getMessage().contains("table 'BudgetTable'"));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet shiftTableSheet = workbook.createSheet("ShiftTableColumns");
      seedTable(shiftTableSheet, workbook);
      IllegalArgumentException shiftTableFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.shiftColumns(shiftTableSheet, new ExcelColumnSpan(1, 1), 1));
      assertTrue(shiftTableFailure.getMessage().contains("table 'BudgetTable'"));
    }
  }

  @Test
  void columnEditsIgnoreBlankDefinedNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");
      setString(sheet, "B1", "Value");
      workbook.getCTWorkbook().addNewDefinedNames().addNewDefinedName().setName("PendingBudget");
      workbook.getCTWorkbook().getDefinedNames().getDefinedNameArray(0).setStringValue(" ");

      assertDoesNotThrow(() -> controller.insertColumns(sheet, 1, 1));
      assertEquals(2, controller.lastColumnIndex(sheet));
    }
  }

  @Test
  void ungroupRowsAndColumnsExpandCollapsedBandsBeforeRemovingOutline() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), true);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), true);

      controller.ungroupRows(sheet, new ExcelRowSpan(1, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(1, 3));

      List<WorkbookReadResult.RowLayout> rows = controller.rowLayouts(sheet);
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);

      assertFalse(rows.get(1).hidden());
      assertEquals(0, rows.get(1).outlineLevel());
      assertFalse(columns.get(1).hidden());
      assertEquals(0, columns.get(1).outlineLevel());
    }
  }

  @Test
  void deleteColumnsClearsTrailingCellsAndColumnMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "B1", "Drop-1");
      setString(sheet, "C1", "Drop-2");
      setString(sheet, "D1", "Keep");
      sheet.setColumnWidth(3, 8192);
      sheet.setColumnHidden(3, true);

      controller.deleteColumns(sheet, new ExcelColumnSpan(1, 2));

      assertEquals(1, controller.lastColumnIndex(sheet));
      assertEquals("Left", sheet.getRow(0).getCell(0).getStringCellValue());
      assertEquals("Keep", sheet.getRow(0).getCell(1).getStringCellValue());
      assertNull(sheet.getRow(0).getCell(2));
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(2, columns.size());
      assertFalse(columns.get(0).hidden());
      assertTrue(columns.get(1).hidden());
      assertEquals(32.0d, columns.get(1).widthCharacters());
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
  void insertColumnsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("InsertTableColumnsSafe");
      seedTable(tableSheet, workbook, "InsertColumnsTable");

      assertDoesNotThrow(() -> controller.insertColumns(tableSheet, 2, 1));

      XSSFSheet autofilterSheet = workbook.createSheet("InsertAutofilterColumnsSafe");
      seedSheetAutofilter(autofilterSheet);

      assertDoesNotThrow(() -> controller.insertColumns(autofilterSheet, 2, 1));

      XSSFSheet validationSheet = workbook.createSheet("InsertValidationColumnsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "C1", "Tail");

      assertDoesNotThrow(() -> controller.insertColumns(validationSheet, 1, 1));
    }
  }

  @Test
  void deleteColumnsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("DeleteTableColumnsSafe");
      seedTable(tableSheet, workbook, "DeleteColumnsTable");
      setString(tableSheet, "C1", "Tail");

      assertDoesNotThrow(() -> controller.deleteColumns(tableSheet, new ExcelColumnSpan(2, 2)));

      XSSFSheet autofilterSheet = workbook.createSheet("DeleteAutofilterColumnsSafe");
      seedSheetAutofilter(autofilterSheet);
      setString(autofilterSheet, "C1", "Tail");

      assertDoesNotThrow(
          () -> controller.deleteColumns(autofilterSheet, new ExcelColumnSpan(2, 2)));

      XSSFSheet validationSheet = workbook.createSheet("DeleteValidationColumnsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "C1", "Tail");

      assertDoesNotThrow(
          () -> controller.deleteColumns(validationSheet, new ExcelColumnSpan(2, 2)));
    }
  }

  @Test
  void shiftColumnsAllowsUntouchedTablesAutofiltersAndDataValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet tableSheet = workbook.createSheet("ShiftTableColumnsSafe");
      seedTable(tableSheet, workbook, "ShiftColumnsTable");
      setString(tableSheet, "K1", "Tail");

      assertDoesNotThrow(() -> controller.shiftColumns(tableSheet, new ExcelColumnSpan(10, 10), 1));

      XSSFSheet autofilterSheet = workbook.createSheet("ShiftAutofilterColumnsSafe");
      seedSheetAutofilter(autofilterSheet);
      setString(autofilterSheet, "K1", "Tail");

      assertDoesNotThrow(
          () -> controller.shiftColumns(autofilterSheet, new ExcelColumnSpan(10, 10), 1));

      XSSFSheet validationSheet = workbook.createSheet("ShiftValidationColumnsSafe");
      seedDataValidation(validationSheet);
      setString(validationSheet, "K1", "Tail");

      assertDoesNotThrow(
          () -> controller.shiftColumns(validationSheet, new ExcelColumnSpan(10, 10), 1));
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

  @Test
  void invalidStoredRangesAreRejectedDuringNormalization() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");
      sheet.getCTWorksheet().addNewAutoFilter().setRef("NOT_A_RANGE");

      IllegalArgumentException failure =
          assertThrows(IllegalArgumentException.class, () -> controller.insertRows(sheet, 0, 1));
      assertTrue(failure.getMessage().contains("Stored sheet autofilter range is invalid"));
    }
  }

  @Test
  void groupingAtWorkbookLimitsSkipsControlArtifacts() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet rowLimitSheet = workbook.createSheet("RowLimit");
      rowLimitSheet.createRow(ExcelRowSpan.MAX_ROW_INDEX).createCell(0).setCellValue("Edge");

      assertDoesNotThrow(
          () ->
              controller.groupRows(
                  rowLimitSheet,
                  new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX, ExcelRowSpan.MAX_ROW_INDEX),
                  true));
      assertDoesNotThrow(
          () ->
              controller.ungroupRows(
                  rowLimitSheet,
                  new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX, ExcelRowSpan.MAX_ROW_INDEX)));

      XSSFSheet columnLimitSheet = workbook.createSheet("ColumnLimit");
      setString(columnLimitSheet, "XFD1", "Edge");

      assertDoesNotThrow(
          () ->
              controller.groupColumns(
                  columnLimitSheet,
                  new ExcelColumnSpan(
                      ExcelColumnSpan.MAX_COLUMN_INDEX, ExcelColumnSpan.MAX_COLUMN_INDEX),
                  false));
      assertDoesNotThrow(
          () ->
              controller.ungroupColumns(
                  columnLimitSheet,
                  new ExcelColumnSpan(
                      ExcelColumnSpan.MAX_COLUMN_INDEX, ExcelColumnSpan.MAX_COLUMN_INDEX)));
    }
  }

  @Test
  void insertColumnsPreservesExplicitMetadataBeforeAndAfterInsertionPoint() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "C1", "Right");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(2, 4096);
      sheet.setColumnHidden(2, true);

      controller.insertColumns(sheet, 1, 1);

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(4, columns.size());
      assertEquals(8.0d, columns.get(0).widthCharacters());
      assertEquals(16.0d, columns.get(3).widthCharacters());
      assertTrue(columns.get(3).hidden());
    }
  }

  @Test
  void deleteColumnsPreservesExplicitMetadataBeforeDeletedBandAndAtTail() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet metadataSheet = workbook.createSheet("Budget");
      setString(metadataSheet, "A1", "Left");
      setString(metadataSheet, "D1", "Right");
      metadataSheet.setColumnWidth(0, 2048);
      metadataSheet.setColumnWidth(3, 4096);
      metadataSheet.setColumnHidden(3, true);

      controller.deleteColumns(metadataSheet, new ExcelColumnSpan(1, 1));

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(metadataSheet);
      assertEquals(3, columns.size());
      assertEquals(8.0d, columns.get(0).widthCharacters());
      assertEquals(16.0d, columns.get(2).widthCharacters());
      assertTrue(columns.get(2).hidden());

      XSSFSheet tailSheet = workbook.createSheet("TailDelete");
      setString(tailSheet, "A1", "Left");
      setString(tailSheet, "B1", "Middle");
      setString(tailSheet, "C1", "Tail");

      controller.deleteColumns(tailSheet, new ExcelColumnSpan(2, 2));

      assertEquals(1, controller.lastColumnIndex(tailSheet));
    }
  }

  @Test
  void shiftColumnsPreservesExplicitMetadataOutsideSourceAndDestination() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Before");
      setString(sheet, "C1", "Source-1");
      setString(sheet, "D1", "Source-2");
      setString(sheet, "E1", "Overwrite-1");
      setString(sheet, "F1", "Overwrite-2");
      setString(sheet, "G1", "After");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(2, 3072);
      sheet.setColumnWidth(3, 4096);
      sheet.setColumnWidth(4, 5120);
      sheet.setColumnWidth(5, 6144);
      sheet.setColumnWidth(6, 7168);

      controller.shiftColumns(sheet, new ExcelColumnSpan(2, 3), 2);

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(7, columns.size());
      assertEquals(8.0d, columns.get(0).widthCharacters());
      assertEquals(12.0d, columns.get(4).widthCharacters());
      assertEquals(16.0d, columns.get(5).widthCharacters());
      assertEquals(28.0d, columns.get(6).widthCharacters());
    }
  }

  @Test
  void workbookContainsFormulaDefinedNamesSkipsBlankReferences() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");

      assertFalse(
          ExcelRowColumnStructureController.workbookContainsFormulaDefinedNames(
              workbook, List.of(new DefinedNameStub(null, -1))));
      assertFalse(
          ExcelRowColumnStructureController.workbookContainsFormulaDefinedNames(
              workbook, List.of(new DefinedNameStub(" ", -1))));
      assertTrue(
          ExcelRowColumnStructureController.workbookContainsFormulaDefinedNames(
              workbook, List.of(new DefinedNameStub("OFFSET(Budget!$A$1,0,0,2,1)", -1))));
    }
  }

  @Test
  void resolvedRangeBackedTargetSkipsUnsetBlankAndFormulaDefinedNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");

      assertTrue(
          ExcelRowColumnStructureController.resolvedRangeBackedTarget(
                  workbook, new DefinedNameStub(null, -1))
              .isEmpty());
      assertTrue(
          ExcelRowColumnStructureController.resolvedRangeBackedTarget(
                  workbook, new DefinedNameStub(" ", -1))
              .isEmpty());
      assertTrue(
          ExcelRowColumnStructureController.resolvedRangeBackedTarget(
                  workbook, new DefinedNameStub("OFFSET(Budget!$A$1,0,0,2,1)", -1))
              .isEmpty());
      assertEquals(
          new ExcelNamedRangeTarget("Budget", "A1:B2"),
          ExcelRowColumnStructureController.resolvedRangeBackedTarget(
                  workbook, new DefinedNameStub("Budget!$A$1:$B$2", -1))
              .orElseThrow());
    }
  }

  @Test
  void resolvedRangeBackedNamesFilterOutUnsetBlankAndFormulaDefinedNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");

      List<ExcelRowColumnStructureController.ResolvedNamedRange> resolved =
          ExcelRowColumnStructureController.resolvedRangeBackedNames(
              workbook,
              List.of(
                  new DefinedNameStub(null, -1),
                  new DefinedNameStub(" ", -1),
                  new DefinedNameStub("OFFSET(Budget!$A$1,0,0,2,1)", -1),
                  new DefinedNameStub("Budget!$A$1:$B$2", -1)));

      assertEquals(1, resolved.size());
      assertEquals("TestName", resolved.getFirst().name());
      assertEquals(new ExcelNamedRangeTarget("Budget", "A1:B2"), resolved.getFirst().target());
      assertEquals(new ExcelRange(0, 1, 0, 1), resolved.getFirst().range());
    }
  }

  @Test
  void shiftWouldCorruptRowsDistinguishesMovedPartialDestinationAndSafeRanges() {
    assertFalse(
        ExcelRowColumnStructureController.shiftWouldCorruptRows(
            new ExcelRange(2, 3, 0, 0), new ExcelRowSpan(2, 3), -2));
    assertTrue(
        ExcelRowColumnStructureController.shiftWouldCorruptRows(
            new ExcelRange(2, 4, 0, 0), new ExcelRowSpan(2, 3), -2));
    assertTrue(
        ExcelRowColumnStructureController.shiftWouldCorruptRows(
            new ExcelRange(0, 1, 0, 0), new ExcelRowSpan(2, 3), -2));
    assertFalse(
        ExcelRowColumnStructureController.shiftWouldCorruptRows(
            new ExcelRange(5, 6, 0, 0), new ExcelRowSpan(2, 3), -2));
  }

  @Test
  void shiftWouldCorruptColumnsDistinguishesMovedPartialDestinationAndSafeRanges() {
    assertFalse(
        ExcelRowColumnStructureController.shiftWouldCorruptColumns(
            new ExcelRange(0, 0, 2, 3), new ExcelColumnSpan(2, 3), -2));
    assertTrue(
        ExcelRowColumnStructureController.shiftWouldCorruptColumns(
            new ExcelRange(0, 0, 2, 4), new ExcelColumnSpan(2, 3), -2));
    assertTrue(
        ExcelRowColumnStructureController.shiftWouldCorruptColumns(
            new ExcelRange(0, 0, 0, 1), new ExcelColumnSpan(2, 3), -2));
    assertFalse(
        ExcelRowColumnStructureController.shiftWouldCorruptColumns(
            new ExcelRange(0, 0, 5, 6), new ExcelColumnSpan(2, 3), -2));
  }

  @Test
  void affectsRowsDetectsRangesBeforeAndAfterMovedBand() {
    assertTrue(
        ExcelRowColumnStructureController.affectsRows(
            new ExcelRange(2, 4, 0, 0), new ExcelRowSpan(1, 2), 2));
    assertFalse(
        ExcelRowColumnStructureController.affectsRows(
            new ExcelRange(10, 12, 0, 0), new ExcelRowSpan(0, 0), 1));
    assertFalse(
        ExcelRowColumnStructureController.affectsRows(
            new ExcelRange(0, 1, 0, 0), new ExcelRowSpan(10, 10), 1));
  }

  @Test
  void affectsColumnsDetectsRangesBeforeAndAfterMovedBand() {
    assertTrue(
        ExcelRowColumnStructureController.affectsColumns(
            new ExcelRange(0, 0, 2, 4), new ExcelColumnSpan(1, 2), 2));
    assertFalse(
        ExcelRowColumnStructureController.affectsColumns(
            new ExcelRange(0, 0, 10, 12), new ExcelColumnSpan(0, 0), 1));
    assertFalse(
        ExcelRowColumnStructureController.affectsColumns(
            new ExcelRange(0, 0, 0, 1), new ExcelColumnSpan(10, 10), 1));
  }

  @Test
  void deleteColumnsDropsExplicitMetadataInsideDeletedBand() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "B1", "Drop");
      setString(sheet, "C1", "Keep");
      sheet.setColumnWidth(1, 4096);
      sheet.setColumnWidth(2, 2048);

      controller.deleteColumns(sheet, new ExcelColumnSpan(1, 1));

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(2, columns.size());
      assertEquals(8.0d, columns.get(1).widthCharacters());
    }
  }

  @Test
  void columnStructuralEditsPreserveOnePoiColumnContainerWithoutExplicitMetadata()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet insertSheet = workbook.createSheet("Insert");
      setString(insertSheet, "A1", "Left");
      setString(insertSheet, "B1", "Right");

      controller.insertColumns(insertSheet, 1, 1);

      assertEquals(1, insertSheet.getCTWorksheet().sizeOfColsArray());

      XSSFSheet shiftSheet = workbook.createSheet("Shift");
      setString(shiftSheet, "A1", "Left");
      setString(shiftSheet, "B1", "Middle");
      setString(shiftSheet, "C1", "Right");

      controller.shiftColumns(shiftSheet, new ExcelColumnSpan(0, 1), 1);

      assertEquals(1, shiftSheet.getCTWorksheet().sizeOfColsArray());

      XSSFSheet deleteSheet = workbook.createSheet("Delete");
      setString(deleteSheet, "A1", "Left");
      setString(deleteSheet, "B1", "Drop");
      setString(deleteSheet, "C1", "Keep");

      controller.deleteColumns(deleteSheet, new ExcelColumnSpan(1, 1));

      assertEquals(1, deleteSheet.getCTWorksheet().sizeOfColsArray());
    }
  }

  @Test
  void groupColumnsNormalizesMissingPoiColumnContainerBeforeGrouping() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "B1", "Middle");
      setString(sheet, "C1", "Right");
      sheet.getCTWorksheet().setColsArray(new CTCols[0]);

      controller.groupColumns(sheet, new ExcelColumnSpan(1, 2), false);

      assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(1, columns.get(1).outlineLevel());
      assertEquals(1, columns.get(2).outlineLevel());
      assertFalse(columns.get(1).hidden());
      assertFalse(columns.get(2).hidden());
    }
  }

  @Test
  void groupColumnsStillWorksAfterColumnStructuralEditsWithoutExplicitMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "B1", "Middle");
      setString(sheet, "C1", "Right");

      controller.shiftColumns(sheet, new ExcelColumnSpan(0, 1), 1);
      controller.insertColumns(sheet, 0, 1);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 2), false);

      assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(1, columns.get(1).outlineLevel());
      assertEquals(1, columns.get(2).outlineLevel());
      assertFalse(columns.get(1).hidden());
      assertFalse(columns.get(2).hidden());
    }
  }

  @Test
  void groupColumnsClearsCollapsedStateOnlyOnTheMatchingControlColumn() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "C1", "Control");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(2, 4096);

      controller.groupColumns(sheet, new ExcelColumnSpan(1, 1), false);

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(8.0d, columns.get(0).widthCharacters());
      assertFalse(columns.get(2).collapsed());
    }
  }

  @Test
  void expandedOuterColumnGroupingDoesNotExpandCollapsedDescendantsOrCrash() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.groupColumns(sheet, new ExcelColumnSpan(2, 3), true);
      assertDoesNotThrow(() -> controller.groupColumns(sheet, new ExcelColumnSpan(0, 2), false));

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(5, columns.size(), "columns=" + columns);
      assertFalse(columns.get(0).hidden(), "columns=" + columns);
      assertFalse(columns.get(1).hidden(), "columns=" + columns);
      assertTrue(columns.get(2).hidden(), "columns=" + columns);
      assertEquals(2, columns.get(2).outlineLevel(), "columns=" + columns);
      assertTrue(columns.get(3).hidden(), "columns=" + columns);
      assertEquals(1, columns.get(3).outlineLevel(), "columns=" + columns);
      assertTrue(columns.get(4).collapsed(), "columns=" + columns);
    }
  }

  @Test
  void repeatedExpandedColumnGroupingCanonicalizesOutlineLevelsBeforeRoundTrip() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-column-outline-", ".xlsx");

    try {
      try (XSSFWorkbook workbook = new XSSFWorkbook()) {
        XSSFSheet sheet = workbook.createSheet("Budget");
        ExcelColumnSpan repeatedColumn = new ExcelColumnSpan(2, 2);

        controller.groupColumns(sheet, new ExcelColumnSpan(2, 3), false);
        for (int repetition = 0; repetition < 6; repetition++) {
          controller.groupColumns(sheet, repeatedColumn, false);
        }
        controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), false);

        assertCanonicalColumnDefinitions(sheet);
        assertColumnOutlineLevels(sheet, 1, 7, 2);

        try (var outputStream = Files.newOutputStream(workbookPath)) {
          workbook.write(outputStream);
        }
      }

      try (XSSFWorkbook reopenedWorkbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
        XSSFSheet reopenedSheet = reopenedWorkbook.getSheet("Budget");

        assertCanonicalColumnDefinitions(reopenedSheet);
        assertColumnOutlineLevels(reopenedSheet, 1, 7, 2);
      }
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  @Test
  void insertColumnsNormalizesMultiColumnRangeDefinitionsBeforeShifting() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Live");
      CTCol groupedRange = addRawColumnDefinition(sheet, 0, 1);
      groupedRange.setHidden(true);
      groupedRange.setOutlineLevel((short) 1);

      controller.insertColumns(sheet, 0, 1);

      assertCanonicalColumnDefinitions(sheet);
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(3, columns.size(), "columns=" + columns);
      assertFalse(columns.get(0).hidden(), "columns=" + columns);
      assertTrue(columns.get(1).hidden(), "columns=" + columns);
      assertTrue(columns.get(2).hidden(), "columns=" + columns);
    }
  }

  @Test
  void insertColumnsDropsSemanticallyEmptyDefinitionsBeforeShifting() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Live");
      addRawColumnDefinition(sheet, 0, 0);

      controller.insertColumns(sheet, 0, 1);

      assertCanonicalColumnDefinitions(sheet);
      assertEquals(0, sheet.getCTWorksheet().getColsArray(0).sizeOfColArray());
    }
  }

  @Test
  void insertColumnsNormalizesOverlappingDefinitionsBeforeShifting() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Live");

      CTCol first = addRawColumnDefinition(sheet, 0, 1);
      first.setHidden(true);
      first.setOutlineLevel((short) 1);

      CTCol second = addRawColumnDefinition(sheet, 1, 2);
      second.setOutlineLevel((short) 2);

      controller.insertColumns(sheet, 0, 1);

      assertCanonicalColumnDefinitions(sheet);
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(4, columns.size(), "columns=" + columns);
      assertEquals(1, columns.get(1).outlineLevel(), "columns=" + columns);
      assertEquals(2, columns.get(2).outlineLevel(), "columns=" + columns);
      assertEquals(2, columns.get(3).outlineLevel(), "columns=" + columns);
    }
  }

  @Test
  void insertColumnsNormalizesDuplicateSingleColumnDefinitionsBeforeShifting() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Live");

      CTCol hidden = addRawColumnDefinition(sheet, 0, 0);
      hidden.setHidden(true);
      hidden.setOutlineLevel((short) 1);

      CTCol nested = addRawColumnDefinition(sheet, 0, 0);
      nested.setOutlineLevel((short) 2);

      controller.insertColumns(sheet, 0, 1);

      assertCanonicalColumnDefinitions(sheet);
      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(2, columns.size(), "columns=" + columns);
      assertEquals(2, columns.get(1).outlineLevel(), "columns=" + columns);
      assertTrue(columns.get(1).hidden(), "columns=" + columns);
    }
  }

  @Test
  void canonicalizeColumnDefinitionsPreservesBestFitOnlyColumns() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      CTCol definition = addRawColumnDefinition(sheet, 0, 0);
      definition.setBestFit(true);

      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(sheet);

      CTCol canonical = sheet.getCTWorksheet().getColsArray(0).getColArray(0);
      assertTrue(canonical.getBestFit());
      assertFalse(canonical.getCustomWidth());
    }
  }

  @Test
  void canonicalizeColumnDefinitionsPreservesPhoneticOnlyColumns() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      CTCol definition = addRawColumnDefinition(sheet, 0, 0);
      definition.setPhonetic(true);

      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(sheet);

      CTCol canonical = sheet.getCTWorksheet().getColsArray(0).getColArray(0);
      assertTrue(canonical.getPhonetic());
      assertFalse(canonical.getCollapsed());
    }
  }

  @Test
  void canonicalizeColumnDefinitionsPreservesStyledColumnsAndDropsZeroStyleDefaults()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet styledSheet = workbook.createSheet("Styled");
      CTCol styled = addRawColumnDefinition(styledSheet, 0, 0);
      styled.setStyle(7L);

      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(styledSheet);

      assertEquals(1, styledSheet.getCTWorksheet().getColsArray(0).sizeOfColArray());
      assertEquals(7L, styledSheet.getCTWorksheet().getColsArray(0).getColArray(0).getStyle());

      XSSFSheet zeroStyleSheet = workbook.createSheet("ZeroStyle");
      CTCol zeroStyle = addRawColumnDefinition(zeroStyleSheet, 0, 0);
      zeroStyle.setStyle(0L);

      ExcelRowColumnStructureController.canonicalizeColumnDefinitions(zeroStyleSheet);

      assertEquals(0, zeroStyleSheet.getCTWorksheet().getColsArray(0).sizeOfColArray());
    }
  }

  @Test
  void shortColumnGroupingSequencesAvoidUnexpectedPoiRuntimeFailures() throws Exception {
    List<ColumnGroupingMutation> mutations = columnGroupingMutations();

    for (ColumnGroupingMutation first : mutations) {
      for (ColumnGroupingMutation second : mutations) {
        assertUnexpectedRuntimeFailureAbsent(first, second);
      }
    }
  }

  @Test
  void setColumnCollapsedSkipsEarlierRangesBeforeMatchingTheTargetRange() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "D1", "Right");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(3, 4096);

      ExcelRowColumnStructureController.setColumnCollapsed(sheet, 3, true);

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertFalse(columns.get(0).collapsed());
      assertTrue(columns.get(3).collapsed());
    }
  }

  @Test
  void setColumnCollapsedIgnoresLaterRangesWhenTheTargetComesFirst() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "D1", "Right");
      sheet.setColumnWidth(3, 4096);

      ExcelRowColumnStructureController.setColumnCollapsed(sheet, 0, true);

      List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertFalse(columns.get(3).collapsed());
    }
  }

  private static void seedSupportedScenario(
      XSSFWorkbook workbook, XSSFSheet sheet, boolean includeFormulas) {
    setString(sheet, "A1", "Item");
    setString(sheet, "B1", "Value");
    setString(sheet, "C1", "Status");
    setString(sheet, "D1", "Notes");
    setString(sheet, "E1", "Aux");
    setString(sheet, "F1", "Merge");
    setString(sheet, "G1", "Formula");

    setString(sheet, "A2", "Hosting");
    setNumeric(sheet, "B2", 42);
    setString(sheet, "C2", "Open");
    setString(sheet, "D2", "Alpha");
    setString(sheet, "A3", "Support");
    setNumeric(sheet, "B3", 84);
    setString(sheet, "C3", "Closed");
    setString(sheet, "D3", "Beta");
    setString(sheet, "A4", "Ops");
    setNumeric(sheet, "B4", 168);
    setString(sheet, "C4", "Open");
    setString(sheet, "D4", "Gamma");
    setString(sheet, "A5", "Tail");
    setNumeric(sheet, "B5", 7);

    if (includeFormulas) {
      sheet.getRow(1).createCell(6).setCellFormula("SUM(B2:B4)");
      sheet.getRow(2).createCell(6).setCellFormula("A2&B2");
      setString(sheet, "G12", "Anchor");
      sheet.getRow(11).createCell(7).setCellFormula("SUM(B2:B4)");
      sheet.getRow(11).createCell(8).setCellFormula("A2&B2");
    }

    Name name = workbook.createName();
    name.setNameName("BudgetValues");
    name.setRefersToFormula("Budget!$B$2:$B$4");

    sheet.addMergedRegion(CellRangeAddress.valueOf("E2:F3"));

    SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();
    XSSFConditionalFormattingRule rule =
        (XSSFConditionalFormattingRule)
            conditionalFormatting.createConditionalFormattingRule("B2>100");
    conditionalFormatting.addConditionalFormatting(
        new CellRangeAddress[] {CellRangeAddress.valueOf("B2:B5")}, rule);

    Cell cellWithLink = getOrCreateCell(sheet, "D5");
    XSSFHyperlink hyperlink =
        (XSSFHyperlink) workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
    hyperlink.setAddress("https://example.com/report");
    cellWithLink.setHyperlink(hyperlink);

    XSSFComment comment =
        (XSSFComment)
            sheet
                .createDrawingPatriarch()
                .createCellComment(workbook.getCreationHelper().createClientAnchor());
    comment.setString(workbook.getCreationHelper().createRichTextString("Review"));
    cellWithLink.setCellComment(comment);
  }

  private static void seedTable(XSSFSheet sheet, XSSFWorkbook workbook) {
    seedTable(sheet, workbook, "BudgetTable");
  }

  private static void seedTable(XSSFSheet sheet, XSSFWorkbook workbook, String tableName) {
    setString(sheet, "A1", "Item");
    setString(sheet, "B1", "Value");
    setString(sheet, "A2", "Hosting");
    setNumeric(sheet, "B2", 42);
    setString(sheet, "A3", "Support");
    setNumeric(sheet, "B3", 84);
    XSSFTable table =
        sheet.createTable(new AreaReference("A1:B3", workbook.getSpreadsheetVersion()));
    table.setName(tableName);
    table.setDisplayName(tableName);
  }

  private static void seedNamedRange(
      XSSFWorkbook workbook, String nameName, String refersToFormula) {
    Name name = workbook.createName();
    name.setNameName(nameName);
    name.setRefersToFormula(refersToFormula);
  }

  private static void seedBlankNamedRange(XSSFWorkbook workbook) {
    var definedNames =
        workbook.getCTWorkbook().isSetDefinedNames()
            ? workbook.getCTWorkbook().getDefinedNames()
            : workbook.getCTWorkbook().addNewDefinedNames();
    var blankName = definedNames.addNewDefinedName();
    blankName.setName("BlankBudget");
    blankName.setStringValue(" ");
  }

  private static void seedSheetAutofilter(XSSFSheet sheet) {
    setString(sheet, "A1", "Item");
    setString(sheet, "B1", "Value");
    setString(sheet, "A2", "Hosting");
    setNumeric(sheet, "B2", 42);
    sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B2"));
  }

  private static void seedDataValidation(XSSFSheet sheet) {
    setString(sheet, "A1", "Status");
    setString(sheet, "A2", "Open");
    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint =
        helper.createExplicitListConstraint(new String[] {"Open", "Closed"});
    DataValidation validation =
        helper.createValidation(constraint, new CellRangeAddressList(1, 3, 0, 0));
    sheet.addValidationData(validation);
  }

  private static List<String> hyperlinkAddresses(XSSFSheet sheet) {
    return sheet.getHyperlinkList().stream()
        .map(hyperlink -> hyperlink.getCellRef())
        .sorted()
        .toList();
  }

  private static List<String> commentAddresses(XSSFSheet sheet) {
    List<String> addresses = new java.util.ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellComment() != null) {
          addresses.add(cell.getAddress().formatAsString());
        }
      }
    }
    return List.copyOf(addresses);
  }

  private static void setString(XSSFSheet sheet, String address, String value) {
    getOrCreateCell(sheet, address).setCellValue(value);
  }

  private static void setNumeric(XSSFSheet sheet, String address, double value) {
    getOrCreateCell(sheet, address).setCellValue(value);
  }

  private static Cell getOrCreateCell(XSSFSheet sheet, String address) {
    CellReference reference = new CellReference(address);
    Row row = sheet.getRow(reference.getRow());
    if (row == null) {
      row = sheet.createRow(reference.getRow());
    }
    Cell cell = row.getCell(reference.getCol());
    if (cell == null) {
      cell = row.createCell(reference.getCol());
    }
    return cell;
  }

  private static List<ColumnGroupingMutation> columnGroupingMutations() {
    List<ColumnGroupingMutation> mutations = new java.util.ArrayList<>();
    for (int firstColumnIndex = 0; firstColumnIndex <= 3; firstColumnIndex++) {
      for (int lastColumnIndex = firstColumnIndex; lastColumnIndex <= 3; lastColumnIndex++) {
        mutations.add(new ColumnGroupingMutation("group", firstColumnIndex, lastColumnIndex, true));
        mutations.add(
            new ColumnGroupingMutation("group", firstColumnIndex, lastColumnIndex, false));
        mutations.add(
            new ColumnGroupingMutation("ungroup", firstColumnIndex, lastColumnIndex, false));
      }
    }
    return List.copyOf(mutations);
  }

  private void assertColumnOutlineLevels(XSSFSheet sheet, int first, int second, int third) {
    List<WorkbookReadResult.ColumnLayout> columns = controller.columnLayouts(sheet);

    assertEquals(first, columns.get(1).outlineLevel(), "columns=" + columns);
    assertEquals(second, columns.get(2).outlineLevel(), "columns=" + columns);
    assertEquals(third, columns.get(3).outlineLevel(), "columns=" + columns);
  }

  private static void assertCanonicalColumnDefinitions(XSSFSheet sheet) {
    assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());

    boolean[] seenColumns = new boolean[ExcelColumnSpan.MAX_COLUMN_INDEX + 1];
    for (CTCol col : sheet.getCTWorksheet().getColsArray(0).getColList()) {
      assertEquals(col.getMin(), col.getMax(), "column definitions=" + sheet.getCTWorksheet());
      for (int columnIndex = (int) col.getMin() - 1;
          columnIndex <= (int) col.getMax() - 1;
          columnIndex++) {
        assertFalse(
            seenColumns[columnIndex], "column definitions overlap for index " + columnIndex);
        seenColumns[columnIndex] = true;
      }
    }
  }

  private static CTCol addRawColumnDefinition(
      XSSFSheet sheet, int firstColumnIndex, int lastColumnIndex) {
    CTCols cols =
        sheet.getCTWorksheet().sizeOfColsArray() == 0
            ? sheet.getCTWorksheet().addNewCols()
            : sheet.getCTWorksheet().getColsArray(0);
    CTCol definition = cols.addNewCol();
    definition.setMin(firstColumnIndex + 1L);
    definition.setMax(lastColumnIndex + 1L);
    return definition;
  }

  private record ColumnGroupingMutation(
      String operation, int firstColumnIndex, int lastColumnIndex, boolean collapsed) {
    private void apply(ExcelRowColumnStructureController controller, XSSFSheet sheet) {
      ExcelColumnSpan columns = new ExcelColumnSpan(firstColumnIndex, lastColumnIndex);
      if ("group".equals(operation)) {
        controller.groupColumns(sheet, columns, collapsed);
        return;
      }
      controller.ungroupColumns(sheet, columns);
    }

    @Override
    public String toString() {
      return operation
          + "("
          + firstColumnIndex
          + ","
          + lastColumnIndex
          + ("group".equals(operation) ? "," + collapsed : "")
          + ")";
    }
  }

  private void assertUnexpectedRuntimeFailureAbsent(
      ColumnGroupingMutation first, ColumnGroupingMutation second) {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      try {
        first.apply(controller, sheet);
        second.apply(controller, sheet);
      } catch (IllegalArgumentException expected) {
        // Invalid outline edits are allowed; only unexpected runtime failures should fail.
      } catch (RuntimeException unexpected) {
        fail(first + " -> " + second + " threw unexpected runtime failure", unexpected);
      }
    } catch (IOException ioException) {
      fail("Closing workbook should not fail during grouping regression checks", ioException);
    }
  }

  /** Minimal POI Name stub for direct defined-name predicate tests. */
  private static final class DefinedNameStub implements Name {
    private final String refersToFormula;
    private final int sheetIndex;

    private DefinedNameStub(String refersToFormula, int sheetIndex) {
      this.refersToFormula = refersToFormula;
      this.sheetIndex = sheetIndex;
    }

    @Override
    public String getSheetName() {
      return sheetIndex < 0 ? null : "Budget";
    }

    @Override
    public String getNameName() {
      return "TestName";
    }

    @Override
    public void setNameName(String name) {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getRefersToFormula() {
      return refersToFormula;
    }

    @Override
    public void setRefersToFormula(String formulaText) {
      throw new UnsupportedOperationException();
    }

    @Override
    public boolean isFunctionName() {
      return false;
    }

    @Override
    public boolean isDeleted() {
      return false;
    }

    @Override
    public boolean isHidden() {
      return false;
    }

    @Override
    public void setSheetIndex(int index) {
      throw new UnsupportedOperationException();
    }

    @Override
    public int getSheetIndex() {
      return sheetIndex;
    }

    @Override
    public String getComment() {
      return "";
    }

    @Override
    public void setComment(String comment) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setFunction(boolean value) {
      throw new UnsupportedOperationException();
    }
  }
}
