package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.engine.cells.CellTextLimits;
import dev.erst.gridgrind.engine.formulas.FormulaLimits;
import dev.erst.gridgrind.engine.links.LinkAdapter;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Coverage for the explicit Excel-authored text, formula, hyperlink, and style limit helpers. */
class ExcelLimitsTest {
  @Test
  void rejectsCellTextPastTheExcelCellLimit() {
    String oversizedText = "x".repeat(Limit.MAX_CELL_TEXT_LENGTH + 1);

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> CellTextLimits.requireSupportedLength(oversizedText, "text"));

    org.junit.jupiter.api.Assertions.assertTrue(
        failure.getMessage().contains("Excel cell text limit"));
  }

  @Test
  void rejectsNewStylesOnceTheWorkbookStyleCapIsReached() {
    assertDoesNotThrow(
        () -> StyleAdapter.requireCellStyleCapacity(StyleAdapter.MAX_CELL_STYLES - 1));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> StyleAdapter.requireCellStyleCapacity(StyleAdapter.MAX_CELL_STYLES));

    org.junit.jupiter.api.Assertions.assertTrue(
        failure.getMessage().contains("workbook style limit"));
  }

  @Test
  void rejectsNewHyperlinksOnceTheWorksheetCapIsReached() {
    assertDoesNotThrow(
        () -> LinkAdapter.requireWorksheetHyperlinkCapacity(Limit.MAX_HYPERLINKS - 1));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> LinkAdapter.requireWorksheetHyperlinkCapacity(Limit.MAX_HYPERLINKS));

    org.junit.jupiter.api.Assertions.assertTrue(
        failure.getMessage().contains("worksheet hyperlink limit"));
  }

  @Test
  void rejectsFormulasPastTheExcelLengthCeiling() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      String oversizedFormula = "A".repeat(Limit.MAX_FORMULA_LENGTH + 1);

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  FormulaLimits.requireSupportedFormula(
                      new FormulaLimits.CellContext(workbook, 0), oversizedFormula));

      org.junit.jupiter.api.Assertions.assertTrue(
          failure.getMessage().contains("formula length limit"));
      org.junit.jupiter.api.Assertions.assertNotNull(cell);
    }
  }

  @Test
  void rejectsFormulasPastTheNestedFunctionLimit() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      String formula = nestedAbsFormula(Limit.MAX_NESTED_FUNCTION_LEVELS + 1);

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  FormulaLimits.requireSupportedFormula(
                      new FormulaLimits.CellContext(workbook, 0), formula));

      org.junit.jupiter.api.Assertions.assertTrue(
          failure.getMessage().contains("formula nesting limit"));
    }
  }

  @Test
  void rejectsFormulasPastTheFunctionArgumentLimit() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      String formula = "SUM(" + "1,".repeat(Limit.MAX_FUNCTION_ARGUMENTS) + "1)";

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  FormulaLimits.requireSupportedFormula(
                      new FormulaLimits.CellContext(workbook, 0), formula));

      org.junit.jupiter.api.Assertions.assertTrue(
          failure.getMessage().contains("function argument limit"));
    }
  }

  @Test
  void acceptsSupportedFormulaShapesWithQuotedTextEscapesAndGroupingParens() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");

      assertDoesNotThrow(
          () ->
              FormulaLimits.requireSupportedFormula(
                  new FormulaLimits.CellContext(workbook, 0),
                  "CONCAT(\"A\"\"B\",TEXT((1+2),\"0\"))"));
    }
  }

  @Test
  void classifiesFormulaShapeLexicallyAcrossFunctionNameVariants() {
    FormulaLimits.FormulaShape shape =
        FormulaLimits.scanFormulaShape("SUM(1,ABS(2),_XLFN.TEXTAFTER(\"a,b\",\",\"))");

    assertEquals(2, shape.maximumFunctionNesting());
    assertEquals(3, shape.maximumFunctionArguments());
    assertTrue(FormulaLimits.isFunctionIdentifierCharacter('A'));
    assertTrue(FormulaLimits.isFunctionIdentifierCharacter('_'));
    assertTrue(FormulaLimits.isFunctionIdentifierCharacter('.'));
    assertFalse(FormulaLimits.isFunctionIdentifierCharacter('('));
  }

  @Test
  void recognizesWhitespaceSeparatedFunctionCallsButRejectsGroupingParens() {
    assertTrue(FormulaLimits.looksLikeFunctionCall("ABS (1)", 4));
    assertFalse(FormulaLimits.looksLikeFunctionCall("(1+2)", 0));
    assertFalse(FormulaLimits.looksLikeFunctionCall("A1 + (1+2)", 5));
  }

  @Test
  void treatsWhitespaceOnlyBodiesOuterCommasAndGroupingParensAsNonArguments() {
    FormulaLimits.FormulaShape whitespaceOnlyFunction = FormulaLimits.scanFormulaShape("SUM(   )");
    FormulaLimits.FormulaShape whitespaceSeparatedFunction =
        FormulaLimits.scanFormulaShape("ABS (1)");
    FormulaLimits.FormulaShape groupingOnly = FormulaLimits.scanFormulaShape("(1+2)");
    FormulaLimits.FormulaShape outerCommaOnly = FormulaLimits.scanFormulaShape("1,2");

    assertEquals(1, whitespaceOnlyFunction.maximumFunctionNesting());
    assertEquals(0, whitespaceOnlyFunction.maximumFunctionArguments());
    assertEquals(1, whitespaceSeparatedFunction.maximumFunctionNesting());
    assertEquals(1, whitespaceSeparatedFunction.maximumFunctionArguments());
    assertEquals(0, groupingOnly.maximumFunctionNesting());
    assertEquals(0, groupingOnly.maximumFunctionArguments());
    assertEquals(0, outerCommaOnly.maximumFunctionNesting());
    assertEquals(0, outerCommaOnly.maximumFunctionArguments());
  }

  @Test
  void rejectsUnsupportedWorkbookContextsAndNegativeSheetIndexes() throws Exception {
    try (HSSFWorkbook workbook = new HSSFWorkbook()) {
      workbook.createSheet("Legacy");

      IllegalArgumentException workbookFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  FormulaLimits.requireSupportedFormula(
                      new FormulaLimits.CellContext(workbook, 0), "SUM(1,2)"));
      assertTrue(workbookFailure.getMessage().contains("XSSF workbook-backed"));
    }

    NullPointerException nullWorkbookFailure =
        assertThrows(NullPointerException.class, () -> new FormulaLimits.CellContext(null, 0));
    assertEquals("workbook must not be null", nullWorkbookFailure.getMessage());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      IllegalArgumentException sheetIndexFailure =
          assertThrows(
              IllegalArgumentException.class, () -> new FormulaLimits.CellContext(workbook, -1));
      assertTrue(sheetIndexFailure.getMessage().contains("sheetIndex must not be negative"));
    }
  }

  private static String nestedAbsFormula(int depth) {
    StringBuilder formula = new StringBuilder("1");
    for (int index = 0; index < depth; index++) {
      formula.insert(0, "ABS(").append(')');
    }
    return formula.toString();
  }
}
