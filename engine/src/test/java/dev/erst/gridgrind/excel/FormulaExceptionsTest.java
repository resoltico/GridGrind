package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for FormulaExceptions POI exception translation. */
class FormulaExceptionsTest {
  @Test
  void classifiesFormulaFailuresDeterministically() {
    RuntimeException invalid =
        FormulaExceptions.wrap(
            "Budget",
            "B4",
            "SUM(",
            new org.apache.poi.ss.formula.FakeFormulaFailure("bad formula"));
    assertInstanceOf(InvalidFormulaException.class, invalid);
    assertEquals("Budget", ((InvalidFormulaException) invalid).sheetName());
    assertEquals("B4", ((InvalidFormulaException) invalid).address());
    assertEquals("SUM(", ((InvalidFormulaException) invalid).formula());
    assertEquals("Invalid formula at Budget!B4: SUM(", invalid.getMessage());

    RuntimeException unsupported =
        FormulaExceptions.wrap(
            "Budget",
            "C4",
            "LAMBDA(x,x+1)(2)",
            new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException("unsupported"));
    assertInstanceOf(UnsupportedFormulaException.class, unsupported);
    assertEquals("Budget", ((UnsupportedFormulaException) unsupported).sheetName());
    assertEquals("C4", ((UnsupportedFormulaException) unsupported).address());
    assertEquals("LAMBDA(x,x+1)(2)", ((UnsupportedFormulaException) unsupported).formula());
    assertEquals(
        "Unsupported formula function LAMBDA at Budget!C4: LAMBDA(x,x+1)(2)",
        unsupported.getMessage());

    RuntimeException nullMessage =
        FormulaExceptions.wrap(
            "Budget", "A1", "A1", new org.apache.poi.ss.formula.FakeFormulaFailure(null));
    assertEquals("Invalid formula at Budget!A1: A1", nullMessage.getMessage());
  }

  @Test
  void leavesUnclassifiedRuntimeExceptionsUntouched() {
    RuntimeException original = new RuntimeException("boom");
    RuntimeException evalFailure = new org.apache.poi.ss.formula.eval.FakeEvalFailure("eval boom");

    assertSame(original, FormulaExceptions.wrap("Budget", "A1", "A1", original));
    assertSame(evalFailure, FormulaExceptions.wrap("Budget", "A1", "A1", evalFailure));
  }

  @Test
  void extractsFunctionNamesConservatively() {
    assertNull(FormulaExceptions.leadingFunctionName(null));
    assertNull(FormulaExceptions.leadingFunctionName("A1"));
    assertNull(FormulaExceptions.leadingFunctionName(" (A1)"));
    assertNull(FormulaExceptions.leadingFunctionName("1SUM(A1)"));
    assertEquals("TEXTAFTER", FormulaExceptions.leadingFunctionName("TEXTAFTER(\"a,b\",\",\")"));
    assertEquals(
        "_XLFN.TEXTAFTER", FormulaExceptions.leadingFunctionName("_XLFN.TEXTAFTER(\"a,b\",\",\")"));
  }
}
