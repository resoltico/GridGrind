package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/** Column structural edit and guard coverage. */
class ExcelColumnStructureEditCoverageTest extends ExcelRowColumnStructureTestSupport {
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
      assertDoesNotThrow(() -> controller.insertRows(insertValidationSheet, 1, 1));
      assertEquals(List.of("A3:A5"), dataValidationRanges(insertValidationSheet));

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
      assertDoesNotThrow(() -> controller.insertColumns(insertValidationSheet, 0, 1));
      assertEquals(List.of("B2:B4"), dataValidationRanges(insertValidationSheet));

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
  void insertColumnsPreservesExplicitMetadataBeforeAndAfterInsertionPoint() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "C1", "Right");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(2, 4096);
      sheet.setColumnHidden(2, true);

      controller.insertColumns(sheet, 1, 1);

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
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

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(metadataSheet);
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

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(7, columns.size());
      assertEquals(8.0d, columns.get(0).widthCharacters());
      assertEquals(12.0d, columns.get(4).widthCharacters());
      assertEquals(16.0d, columns.get(5).widthCharacters());
      assertEquals(28.0d, columns.get(6).widthCharacters());
    }
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

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(2, columns.size());
      assertEquals(8.0d, columns.get(1).widthCharacters());
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
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(3, columns.size(), "columns=" + columns);
      assertTrue(columns.get(0).hidden(), "columns=" + columns);
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
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
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
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
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
}
