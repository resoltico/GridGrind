package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end coverage for POI-backed array-formula authoring and readback. */
class ExcelArrayFormulaTest {
  @Test
  void arrayFormulaGroupsRoundTripThroughDirectSheetApis() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-array-formula-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Calc");
      seedSourceData(sheet);

      sheet.setArrayFormula("D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4"));
      sheet.setArrayFormula("F2", new ExcelArrayFormulaDefinition("SUM(B2:C2)"));

      ExcelArrayFormulaSnapshot multiCell =
          sheet.arrayFormulas().stream()
              .filter(snapshot -> "D2:D4".equals(snapshot.range()))
              .findFirst()
              .orElseThrow();
      ExcelArrayFormulaSnapshot singleCell =
          sheet.arrayFormulas().stream()
              .filter(snapshot -> "F2".equals(snapshot.range()))
              .findFirst()
              .orElseThrow();
      assertEquals("D2", multiCell.topLeftAddress());
      assertEquals("B2:B4*C2:C4", multiCell.formula());
      assertFalse(multiCell.singleCell());
      assertTrue(singleCell.singleCell());

      ExcelCellSnapshot.FormulaSnapshot formulaCell =
          assertInstanceOf(ExcelCellSnapshot.FormulaSnapshot.class, sheet.snapshotCell("D3"));
      assertEquals("B2:B4*C2:C4", formulaCell.formula());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = reopened.sheet("Calc");
      assertEquals(2, sheet.arrayFormulas().size());
      assertEquals(
          List.of("D2:D4", "F2"),
          sheet.arrayFormulas().stream().map(ExcelArrayFormulaSnapshot::range).sorted().toList());

      sheet.clearArrayFormula("D3");
      assertEquals(
          List.of("F2"),
          sheet.arrayFormulas().stream().map(ExcelArrayFormulaSnapshot::range).toList());
      assertEquals("BLANK", sheet.snapshotCell("D2").effectiveType());

      IllegalArgumentException notArray =
          assertThrows(IllegalArgumentException.class, () -> sheet.clearArrayFormula("A1"));
      assertTrue(notArray.getMessage().contains("is not part of an array formula"));
    }
  }

  @Test
  void arrayFormulaGroupsAreVisibleThroughWorkbookIntrospection() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      ExcelSheet sheet = workbook.getOrCreateSheet("Calc");
      seedSourceData(sheet);
      sheet.setArrayFormula("D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4"));

      WorkbookSheetResult.ArrayFormulasResult result =
          assertInstanceOf(
              WorkbookSheetResult.ArrayFormulasResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetArrayFormulas(
                      "array-formulas", new ExcelSheetSelection.Selected(List.of("Calc")))));
      assertEquals("array-formulas", result.stepId());
      assertEquals(1, result.arrayFormulas().size());
      assertEquals("Calc", result.arrayFormulas().getFirst().sheetName());
      assertEquals("D2:D4", result.arrayFormulas().getFirst().range());
    }
  }

  private static void seedSourceData(ExcelSheet sheet) {
    sheet.setCell("A1", ExcelCellValue.text("Month"));
    sheet.setCell("B1", ExcelCellValue.text("Plan"));
    sheet.setCell("C1", ExcelCellValue.text("Actual"));
    sheet.setCell("A2", ExcelCellValue.text("Jan"));
    sheet.setCell("A3", ExcelCellValue.text("Feb"));
    sheet.setCell("A4", ExcelCellValue.text("Mar"));
    sheet.setCell("B2", ExcelCellValue.number(10d));
    sheet.setCell("B3", ExcelCellValue.number(18d));
    sheet.setCell("B4", ExcelCellValue.number(15d));
    sheet.setCell("C2", ExcelCellValue.number(12d));
    sheet.setCell("C3", ExcelCellValue.number(16d));
    sheet.setCell("C4", ExcelCellValue.number(21d));
  }
}
