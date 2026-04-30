package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.reflect.Method;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Coverage for the explicit Excel-authored text, formula, hyperlink, and style limit helpers. */
class ExcelLimitsTest {
  @Test
  void rejectsCellTextPastTheExcelCellLimit() {
    String oversizedText = "x".repeat(ExcelCellTextLimits.MAX_CELL_TEXT_LENGTH + 1);

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelCellTextLimits.requireSupportedLength(oversizedText, "text"));

    org.junit.jupiter.api.Assertions.assertTrue(
        failure.getMessage().contains("Excel cell text limit"));
  }

  @Test
  void rejectsNewStylesOnceTheWorkbookStyleCapIsReached() {
    assertDoesNotThrow(
        () ->
            ExcelWorkbookStyleLimits.requireCellStyleCapacity(
                ExcelWorkbookStyleLimits.MAX_CELL_STYLES - 1));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelWorkbookStyleLimits.requireCellStyleCapacity(
                    ExcelWorkbookStyleLimits.MAX_CELL_STYLES));

    org.junit.jupiter.api.Assertions.assertTrue(
        failure.getMessage().contains("workbook style limit"));
  }

  @Test
  void rejectsNewHyperlinksOnceTheWorksheetCapIsReached() {
    assertDoesNotThrow(
        () ->
            ExcelHyperlinkLimits.requireWorksheetHyperlinkCapacity(
                ExcelHyperlinkLimits.MAX_HYPERLINKS_PER_SHEET - 1));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelHyperlinkLimits.requireWorksheetHyperlinkCapacity(
                    ExcelHyperlinkLimits.MAX_HYPERLINKS_PER_SHEET));

    org.junit.jupiter.api.Assertions.assertTrue(
        failure.getMessage().contains("worksheet hyperlink limit"));
  }

  @Test
  void rejectsFormulasPastTheExcelLengthCeiling() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      String oversizedFormula = "A".repeat(ExcelFormulaLimits.MAX_FORMULA_LENGTH + 1);

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelFormulaLimits.requireSupportedFormula(
                      new ExcelFormulaLimits.CellContext(workbook, 0), oversizedFormula));

      org.junit.jupiter.api.Assertions.assertTrue(
          failure.getMessage().contains("formula length limit"));
      org.junit.jupiter.api.Assertions.assertNotNull(cell);
    }
  }

  @Test
  void rejectsFormulasPastTheNestedFunctionLimit() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      String formula = nestedAbsFormula(ExcelFormulaLimits.MAX_NESTED_FUNCTION_LEVELS + 1);

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelFormulaLimits.requireSupportedFormula(
                      new ExcelFormulaLimits.CellContext(workbook, 0), formula));

      org.junit.jupiter.api.Assertions.assertTrue(
          failure.getMessage().contains("formula nesting limit"));
    }
  }

  @Test
  void rejectsFormulasPastTheFunctionArgumentLimit() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      String formula = "SUM(" + "1,".repeat(ExcelFormulaLimits.MAX_FUNCTION_ARGUMENTS) + "1)";

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelFormulaLimits.requireSupportedFormula(
                      new ExcelFormulaLimits.CellContext(workbook, 0), formula));

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
              ExcelFormulaLimits.requireSupportedFormula(
                  new ExcelFormulaLimits.CellContext(workbook, 0),
                  "CONCAT(\"A\"\"B\",TEXT((1+2),\"0\"))"));
    }
  }

  @Test
  void classifiesFormulaShapeLexicallyAcrossFunctionNameVariants() throws Throwable {
    Object shape = scanFormulaShape("SUM(1,ABS(2),_XLFN.TEXTAFTER(\"a,b\",\",\"))");

    assertEquals(2, shapeMetric(shape, "maximumFunctionNesting"));
    assertEquals(3, shapeMetric(shape, "maximumFunctionArguments"));
    assertTrue(isFunctionIdentifierCharacter('A'));
    assertTrue(isFunctionIdentifierCharacter('_'));
    assertTrue(isFunctionIdentifierCharacter('.'));
    assertFalse(isFunctionIdentifierCharacter('('));
  }

  @Test
  void recognizesWhitespaceSeparatedFunctionCallsButRejectsGroupingParens() throws Throwable {
    assertTrue(looksLikeFunctionCall("ABS (1)", 4));
    assertFalse(looksLikeFunctionCall("(1+2)", 0));
    assertFalse(looksLikeFunctionCall("A1 + (1+2)", 5));
  }

  @Test
  void treatsWhitespaceOnlyBodiesOuterCommasAndGroupingParensAsNonArguments() throws Throwable {
    Object whitespaceOnlyFunction = scanFormulaShape("SUM(   )");
    Object whitespaceSeparatedFunction = scanFormulaShape("ABS (1)");
    Object groupingOnly = scanFormulaShape("(1+2)");
    Object outerCommaOnly = scanFormulaShape("1,2");

    assertEquals(1, shapeMetric(whitespaceOnlyFunction, "maximumFunctionNesting"));
    assertEquals(0, shapeMetric(whitespaceOnlyFunction, "maximumFunctionArguments"));
    assertEquals(1, shapeMetric(whitespaceSeparatedFunction, "maximumFunctionNesting"));
    assertEquals(1, shapeMetric(whitespaceSeparatedFunction, "maximumFunctionArguments"));
    assertEquals(0, shapeMetric(groupingOnly, "maximumFunctionNesting"));
    assertEquals(0, shapeMetric(groupingOnly, "maximumFunctionArguments"));
    assertEquals(0, shapeMetric(outerCommaOnly, "maximumFunctionNesting"));
    assertEquals(0, shapeMetric(outerCommaOnly, "maximumFunctionArguments"));
  }

  @Test
  void rejectsUnsupportedWorkbookContextsAndNegativeSheetIndexes() throws Exception {
    try (HSSFWorkbook workbook = new HSSFWorkbook()) {
      workbook.createSheet("Legacy");

      IllegalArgumentException workbookFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelFormulaLimits.requireSupportedFormula(
                      new ExcelFormulaLimits.CellContext(workbook, 0), "SUM(1,2)"));
      assertTrue(workbookFailure.getMessage().contains("XSSF workbook-backed"));
    }

    NullPointerException nullWorkbookFailure =
        assertThrows(NullPointerException.class, () -> new ExcelFormulaLimits.CellContext(null, 0));
    assertEquals("workbook must not be null", nullWorkbookFailure.getMessage());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      IllegalArgumentException sheetIndexFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> new ExcelFormulaLimits.CellContext(workbook, -1));
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

  private static Object scanFormulaShape(String formula) throws Throwable {
    Method method = ExcelFormulaLimits.class.getDeclaredMethod("scanFormulaShape", String.class);
    MethodHandles.Lookup lookup =
        MethodHandles.privateLookupIn(ExcelFormulaLimits.class, MethodHandles.lookup());
    return lookup.unreflect(method).invoke(formula);
  }

  private static int shapeMetric(Object shape, String accessorName) throws Throwable {
    MethodHandles.Lookup lookup =
        MethodHandles.privateLookupIn(shape.getClass(), MethodHandles.lookup());
    return (int)
        lookup
            .findVirtual(shape.getClass(), accessorName, MethodType.methodType(int.class))
            .invoke(shape);
  }

  private static boolean looksLikeFunctionCall(String formula, int openParenIndex)
      throws Throwable {
    Method method =
        ExcelFormulaLimits.class.getDeclaredMethod(
            "looksLikeFunctionCall", String.class, int.class);
    MethodHandles.Lookup lookup =
        MethodHandles.privateLookupIn(ExcelFormulaLimits.class, MethodHandles.lookup());
    return (boolean) lookup.unreflect(method).invoke(formula, openParenIndex);
  }

  private static boolean isFunctionIdentifierCharacter(char value) throws Throwable {
    Method method =
        ExcelFormulaLimits.class.getDeclaredMethod("isFunctionIdentifierCharacter", char.class);
    MethodHandles.Lookup lookup =
        MethodHandles.privateLookupIn(ExcelFormulaLimits.class, MethodHandles.lookup());
    return (boolean) lookup.unreflect(method).invoke(value);
  }
}
