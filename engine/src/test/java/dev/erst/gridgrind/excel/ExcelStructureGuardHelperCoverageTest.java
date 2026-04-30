package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Private guard and range-backed helper coverage. */
class ExcelStructureGuardHelperCoverageTest extends ExcelRowColumnStructureTestSupport {
  @Test
  void privateStructureGuardsRejectEachUnsupportedRowAndColumnSurface() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet insertRowTableSheet = workbook.createSheet("InsertRowTable");
      seedTable(insertRowTableSheet, workbook, "InsertRowTable");
      assertTrue(
          unsupportedStructure(
                  () -> controller.rejectAffectedRowStructuresForInsert(insertRowTableSheet, 1))
              .getMessage()
              .contains("table 'InsertRowTable'"));

      XSSFSheet insertRowAutofilterSheet = workbook.createSheet("InsertRowAutofilter");
      seedSheetAutofilter(insertRowAutofilterSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForInsert(insertRowAutofilterSheet, 1))
              .getMessage()
              .contains("sheet autofilter"));

      XSSFSheet insertRowValidationSheet = workbook.createSheet("InsertRowValidation");
      seedDataValidation(insertRowValidationSheet);
      assertDoesNotThrow(
          () -> controller.rejectAffectedRowStructuresForInsert(insertRowValidationSheet, 1));

      XSSFSheet deleteRowTableSheet = workbook.createSheet("DeleteRowTable");
      seedTable(deleteRowTableSheet, workbook, "DeleteRowTable");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForDelete(
                          deleteRowTableSheet, new ExcelRowSpan(1, 1)))
              .getMessage()
              .contains("table 'DeleteRowTable'"));

      XSSFSheet deleteRowAutofilterSheet = workbook.createSheet("DeleteRowAutofilter");
      seedSheetAutofilter(deleteRowAutofilterSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForDelete(
                          deleteRowAutofilterSheet, new ExcelRowSpan(1, 1)))
              .getMessage()
              .contains("sheet autofilter"));

      XSSFSheet deleteRowValidationSheet = workbook.createSheet("DeleteRowValidation");
      seedDataValidation(deleteRowValidationSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForDelete(
                          deleteRowValidationSheet, new ExcelRowSpan(1, 1)))
              .getMessage()
              .contains("data validation"));

      XSSFSheet shiftRowTableSheet = workbook.createSheet("ShiftRowTable");
      seedTable(shiftRowTableSheet, workbook, "ShiftRowTable");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForShift(
                          shiftRowTableSheet, new ExcelRowSpan(1, 1), 1))
              .getMessage()
              .contains("table 'ShiftRowTable'"));

      XSSFSheet shiftRowAutofilterSheet = workbook.createSheet("ShiftRowAutofilter");
      seedSheetAutofilter(shiftRowAutofilterSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForShift(
                          shiftRowAutofilterSheet, new ExcelRowSpan(1, 1), 1))
              .getMessage()
              .contains("sheet autofilter"));

      XSSFSheet shiftRowValidationSheet = workbook.createSheet("ShiftRowValidation");
      seedDataValidation(shiftRowValidationSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedRowStructuresForShift(
                          shiftRowValidationSheet, new ExcelRowSpan(1, 1), 1))
              .getMessage()
              .contains("data validation"));

      XSSFSheet insertColumnTableSheet = workbook.createSheet("InsertColumnTable");
      seedTable(insertColumnTableSheet, workbook, "InsertColumnTable");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForInsert(insertColumnTableSheet, 1))
              .getMessage()
              .contains("table 'InsertColumnTable'"));

      XSSFSheet insertColumnAutofilterSheet = workbook.createSheet("InsertColumnAutofilter");
      seedSheetAutofilter(insertColumnAutofilterSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForInsert(
                          insertColumnAutofilterSheet, 1))
              .getMessage()
              .contains("sheet autofilter"));

      XSSFSheet insertColumnValidationSheet = workbook.createSheet("InsertColumnValidation");
      seedDataValidation(insertColumnValidationSheet);
      assertDoesNotThrow(
          () -> controller.rejectAffectedColumnStructuresForInsert(insertColumnValidationSheet, 0));

      XSSFSheet deleteColumnTableSheet = workbook.createSheet("DeleteColumnTable");
      seedTable(deleteColumnTableSheet, workbook, "DeleteColumnTable");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForDelete(
                          deleteColumnTableSheet, new ExcelColumnSpan(1, 1)))
              .getMessage()
              .contains("table 'DeleteColumnTable'"));

      XSSFSheet deleteColumnAutofilterSheet = workbook.createSheet("DeleteColumnAutofilter");
      seedSheetAutofilter(deleteColumnAutofilterSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForDelete(
                          deleteColumnAutofilterSheet, new ExcelColumnSpan(1, 1)))
              .getMessage()
              .contains("sheet autofilter"));

      XSSFSheet deleteColumnValidationSheet = workbook.createSheet("DeleteColumnValidation");
      seedDataValidation(deleteColumnValidationSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForDelete(
                          deleteColumnValidationSheet, new ExcelColumnSpan(0, 0)))
              .getMessage()
              .contains("data validation"));

      XSSFSheet shiftColumnTableSheet = workbook.createSheet("ShiftColumnTable");
      seedTable(shiftColumnTableSheet, workbook, "ShiftColumnTable");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForShift(
                          shiftColumnTableSheet, new ExcelColumnSpan(1, 1), 1))
              .getMessage()
              .contains("table 'ShiftColumnTable'"));

      XSSFSheet shiftColumnAutofilterSheet = workbook.createSheet("ShiftColumnAutofilter");
      seedSheetAutofilter(shiftColumnAutofilterSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForShift(
                          shiftColumnAutofilterSheet, new ExcelColumnSpan(0, 0), 1))
              .getMessage()
              .contains("sheet autofilter"));

      XSSFSheet shiftColumnValidationSheet = workbook.createSheet("ShiftColumnValidation");
      seedDataValidation(shiftColumnValidationSheet);
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectAffectedColumnStructuresForShift(
                          shiftColumnValidationSheet, new ExcelColumnSpan(0, 0), 1))
              .getMessage()
              .contains("data validation"));
    }
  }

  @Test
  void privateNamedRangeGuardsRejectDestructiveRowAndColumnEdits() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet deleteRowsSheet = workbook.createSheet("DeleteRows");
      setString(deleteRowsSheet, "A3", "Low");
      setString(deleteRowsSheet, "B4", "High");
      seedNamedRange(workbook, "DeleteRowsRange", "DeleteRows!$A$3:$B$4");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectDestructiveNamedRangesForRowDelete(
                          workbook, deleteRowsSheet, new ExcelRowSpan(2, 2)))
              .getMessage()
              .contains("named range 'DeleteRowsRange'"));

      XSSFSheet shiftRowsSheet = workbook.createSheet("ShiftRows");
      setString(shiftRowsSheet, "A1", "Named");
      setString(shiftRowsSheet, "B2", "Range");
      setString(shiftRowsSheet, "A3", "Shifted");
      setString(shiftRowsSheet, "A4", "Rows");
      seedNamedRange(workbook, "ShiftRowsRange", "ShiftRows!$A$1:$B$2");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectDestructiveNamedRangesForRowShift(
                          workbook, shiftRowsSheet, new ExcelRowSpan(2, 3), -2))
              .getMessage()
              .contains("named range 'ShiftRowsRange'"));

      XSSFSheet deleteColumnsSheet = workbook.createSheet("DeleteColumns");
      setString(deleteColumnsSheet, "C1", "Low");
      setString(deleteColumnsSheet, "D2", "High");
      seedNamedRange(workbook, "DeleteColumnsRange", "DeleteColumns!$C$1:$D$2");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectDestructiveNamedRangesForColumnDelete(
                          workbook, deleteColumnsSheet, new ExcelColumnSpan(2, 2)))
              .getMessage()
              .contains("named range 'DeleteColumnsRange'"));

      XSSFSheet shiftColumnsSheet = workbook.createSheet("ShiftColumns");
      setString(shiftColumnsSheet, "A1", "Named");
      setString(shiftColumnsSheet, "B2", "Range");
      setString(shiftColumnsSheet, "C1", "Shifted");
      setString(shiftColumnsSheet, "D1", "Columns");
      seedNamedRange(workbook, "ShiftColumnsRange", "ShiftColumns!$A$1:$B$2");
      assertTrue(
          unsupportedStructure(
                  () ->
                      controller.rejectDestructiveNamedRangesForColumnShift(
                          workbook, shiftColumnsSheet, new ExcelColumnSpan(2, 3), -2))
              .getMessage()
              .contains("named range 'ShiftColumnsRange'"));
    }
  }

  @Test
  void guardDelegatorsAlsoCoverNonDestructiveReturnPaths() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet safeRowSheet = workbook.createSheet("SafeRows");
      setString(safeRowSheet, "A1", "Header");
      setString(safeRowSheet, "A2", "Value");
      assertDoesNotThrow(
          () ->
              controller.rejectAffectedRowStructuresForDelete(
                  safeRowSheet, new ExcelRowSpan(0, 0)));
      assertDoesNotThrow(
          () ->
              controller.rejectAffectedRowStructuresForShift(
                  safeRowSheet, new ExcelRowSpan(0, 0), 1));

      XSSFSheet safeColumnSheet = workbook.createSheet("SafeColumns");
      seedTable(safeColumnSheet, workbook, "SafeTable");
      safeColumnSheet.getCTWorksheet().addNewAutoFilter().setRef("A1:B3");
      assertDoesNotThrow(
          () ->
              controller.rejectAffectedColumnStructuresForDelete(
                  safeColumnSheet, new ExcelColumnSpan(4, 4)));
      assertDoesNotThrow(
          () ->
              controller.rejectAffectedColumnStructuresForShift(
                  safeColumnSheet, new ExcelColumnSpan(4, 4), 1));

      XSSFSheet namesSheet = workbook.createSheet("Names");
      setString(namesSheet, "A1", "Budget");
      setString(namesSheet, "B2", "Values");
      seedNamedRange(workbook, "BudgetValues", "Names!$A$1:$B$2");
      assertDoesNotThrow(
          () ->
              controller.rejectDestructiveNamedRangesForRowDelete(
                  workbook, namesSheet, new ExcelRowSpan(4, 4)));
      assertDoesNotThrow(
          () ->
              controller.rejectDestructiveNamedRangesForRowShift(
                  workbook, namesSheet, new ExcelRowSpan(4, 4), 1));
      assertDoesNotThrow(
          () ->
              controller.rejectDestructiveNamedRangesForColumnDelete(
                  workbook, namesSheet, new ExcelColumnSpan(4, 4)));
      assertDoesNotThrow(
          () ->
              controller.rejectDestructiveNamedRangesForColumnShift(
                  workbook, namesSheet, new ExcelColumnSpan(4, 4), 1));
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
}
