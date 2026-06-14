package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import java.util.function.BiConsumer;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Private guard and range-backed helper coverage. */
class ExcelStructureGuardHelperCoverageTest extends ExcelRowColumnStructureTestSupport {
  @Test
  void privateStructureGuardsRejectEachUnsupportedRowAndColumnSurface() throws Exception {
    for (GuardCase guardCase : guardCases()) {
      assertStructureGuardCase(
          guardCase.sheetName(),
          guardCase.seeder(),
          guardCase.operation(),
          guardCase.allowed(),
          guardCase.expectedMessageFragment());
    }
  }

  @Test
  void privateNamedRangeGuardsRejectDestructiveRowAndColumnEdits() throws Exception {
    assertRowNamedRangeRejected(
        "DeleteRows",
        "DeleteRowsRange",
        "DeleteRows!$A$3:$B$4",
        new ExcelRowSpan(2, 2),
        (workbook, sheet, rows) ->
            runUnchecked(
                () -> controller.rejectDestructiveNamedRangesForRowDelete(workbook, sheet, rows)),
        "A3",
        "Low",
        "B4",
        "High");
    assertRowNamedRangeRejected(
        "ShiftRows",
        "ShiftRowsRange",
        "ShiftRows!$A$1:$B$2",
        new ExcelRowSpan(2, 3),
        (workbook, sheet, rows) ->
            runUnchecked(
                () ->
                    controller.rejectDestructiveNamedRangesForRowShift(workbook, sheet, rows, -2)),
        "A1",
        "Named",
        "B2",
        "Range",
        "A3",
        "Shifted",
        "A4",
        "Rows");
    assertColumnNamedRangeRejected(
        "DeleteColumns",
        "DeleteColumnsRange",
        "DeleteColumns!$C$1:$D$2",
        new ExcelColumnSpan(2, 2),
        (workbook, sheet, columns) ->
            runUnchecked(
                () ->
                    controller.rejectDestructiveNamedRangesForColumnDelete(
                        workbook, sheet, columns)),
        "C1",
        "Low",
        "D2",
        "High");
    assertColumnNamedRangeRejected(
        "ShiftColumns",
        "ShiftColumnsRange",
        "ShiftColumns!$A$1:$B$2",
        new ExcelColumnSpan(2, 3),
        (workbook, sheet, columns) ->
            runUnchecked(
                () ->
                    controller.rejectDestructiveNamedRangesForColumnShift(
                        workbook, sheet, columns, -2)),
        "A1",
        "Named",
        "B2",
        "Range",
        "C1",
        "Shifted",
        "D1",
        "Columns");
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
          assertThrows(IllegalArgumentException.class, () -> rowController.insertRows(sheet, 0, 1));
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
          ExcelNamedRangeTarget.range("Budget", "A1:B2"),
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
      assertEquals(ExcelNamedRangeTarget.range("Budget", "A1:B2"), resolved.getFirst().target());
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

  private List<GuardCase> guardCases() {
    return List.of(
        new GuardCase(
            "InsertRowTable",
            (workbook, sheet) -> seedTable(sheet, workbook, "InsertRowTable"),
            sheet -> runUnchecked(() -> controller.rejectAffectedRowStructuresForInsert(sheet, 1)),
            false,
            "table 'InsertRowTable'"),
        new GuardCase(
            "InsertRowAutofilter",
            (workbook, sheet) -> seedSheetAutofilter(sheet),
            sheet -> runUnchecked(() -> controller.rejectAffectedRowStructuresForInsert(sheet, 1)),
            false,
            "sheet autofilter"),
        new GuardCase(
            "InsertRowValidation",
            (workbook, sheet) -> seedDataValidation(sheet),
            sheet -> runUnchecked(() -> controller.rejectAffectedRowStructuresForInsert(sheet, 1)),
            true,
            ""),
        new GuardCase(
            "DeleteRowTable",
            (workbook, sheet) -> seedTable(sheet, workbook, "DeleteRowTable"),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedRowStructuresForDelete(
                            sheet, new ExcelRowSpan(1, 1))),
            false,
            "table 'DeleteRowTable'"),
        new GuardCase(
            "DeleteRowAutofilter",
            (workbook, sheet) -> seedSheetAutofilter(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedRowStructuresForDelete(
                            sheet, new ExcelRowSpan(1, 1))),
            false,
            "sheet autofilter"),
        new GuardCase(
            "DeleteRowValidation",
            (workbook, sheet) -> seedDataValidation(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedRowStructuresForDelete(
                            sheet, new ExcelRowSpan(1, 1))),
            false,
            "data validation"),
        new GuardCase(
            "ShiftRowTable",
            (workbook, sheet) -> seedTable(sheet, workbook, "ShiftRowTable"),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedRowStructuresForShift(
                            sheet, new ExcelRowSpan(1, 1), 1)),
            false,
            "table 'ShiftRowTable'"),
        new GuardCase(
            "ShiftRowAutofilter",
            (workbook, sheet) -> seedSheetAutofilter(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedRowStructuresForShift(
                            sheet, new ExcelRowSpan(1, 1), 1)),
            false,
            "sheet autofilter"),
        new GuardCase(
            "ShiftRowValidation",
            (workbook, sheet) -> seedDataValidation(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedRowStructuresForShift(
                            sheet, new ExcelRowSpan(1, 1), 1)),
            false,
            "data validation"),
        new GuardCase(
            "InsertColumnTable",
            (workbook, sheet) -> seedTable(sheet, workbook, "InsertColumnTable"),
            sheet ->
                runUnchecked(() -> controller.rejectAffectedColumnStructuresForInsert(sheet, 1)),
            false,
            "table 'InsertColumnTable'"),
        new GuardCase(
            "InsertColumnAutofilter",
            (workbook, sheet) -> seedSheetAutofilter(sheet),
            sheet ->
                runUnchecked(() -> controller.rejectAffectedColumnStructuresForInsert(sheet, 1)),
            false,
            "sheet autofilter"),
        new GuardCase(
            "InsertColumnValidation",
            (workbook, sheet) -> seedDataValidation(sheet),
            sheet ->
                runUnchecked(() -> controller.rejectAffectedColumnStructuresForInsert(sheet, 0)),
            true,
            ""),
        new GuardCase(
            "DeleteColumnTable",
            (workbook, sheet) -> seedTable(sheet, workbook, "DeleteColumnTable"),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedColumnStructuresForDelete(
                            sheet, new ExcelColumnSpan(1, 1))),
            false,
            "table 'DeleteColumnTable'"),
        new GuardCase(
            "DeleteColumnAutofilter",
            (workbook, sheet) -> seedSheetAutofilter(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedColumnStructuresForDelete(
                            sheet, new ExcelColumnSpan(1, 1))),
            false,
            "sheet autofilter"),
        new GuardCase(
            "DeleteColumnValidation",
            (workbook, sheet) -> seedDataValidation(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedColumnStructuresForDelete(
                            sheet, new ExcelColumnSpan(0, 0))),
            false,
            "data validation"),
        new GuardCase(
            "ShiftColumnTable",
            (workbook, sheet) -> seedTable(sheet, workbook, "ShiftColumnTable"),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedColumnStructuresForShift(
                            sheet, new ExcelColumnSpan(1, 1), 1)),
            false,
            "table 'ShiftColumnTable'"),
        new GuardCase(
            "ShiftColumnAutofilter",
            (workbook, sheet) -> seedSheetAutofilter(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedColumnStructuresForShift(
                            sheet, new ExcelColumnSpan(0, 0), 1)),
            false,
            "sheet autofilter"),
        new GuardCase(
            "ShiftColumnValidation",
            (workbook, sheet) -> seedDataValidation(sheet),
            sheet ->
                runUnchecked(
                    () ->
                        controller.rejectAffectedColumnStructuresForShift(
                            sheet, new ExcelColumnSpan(0, 0), 1)),
            false,
            "data validation"));
  }

  private record GuardCase(
      String sheetName,
      BiConsumer<XSSFWorkbook, XSSFSheet> seeder,
      CheckedSheetOperation operation,
      boolean allowed,
      String expectedMessageFragment) {}
}
