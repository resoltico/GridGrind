package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.function.FunctionMetadataRegistry;
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

    RuntimeException missingExternalWorkbook =
        FormulaExceptions.wrap(
            "Budget",
            "D4",
            "[Rates.xlsx]Sheet1!A1",
            new IllegalStateException(
                "outer",
                new FakeWorkbookNotFoundException(
                    "Could not resolve external workbook name 'Rates.xlsx'.")));
    assertInstanceOf(MissingExternalWorkbookException.class, missingExternalWorkbook);
    assertEquals(
        "Missing external workbook Rates.xlsx at Budget!D4: [Rates.xlsx]Sheet1!A1",
        missingExternalWorkbook.getMessage());

    RuntimeException unregisteredUdf =
        FormulaExceptions.wrap(
            "Budget",
            "E4",
            "DOUBLE(A1)",
            new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException("DOUBLE"));
    assertInstanceOf(UnregisteredUserDefinedFunctionException.class, unregisteredUdf);
    assertEquals(
        "User-defined function DOUBLE is not registered at Budget!E4: DOUBLE(A1)",
        unregisteredUdf.getMessage());
  }

  @Test
  void classifiesJdkRuntimeExceptionsWithPoiFormulaStackAsInvalid() {
    IllegalStateException parserFailure =
        new IllegalStateException(
            "Parsed past the end of the formula, pos: 15, length: 14, formula: [^owe_e`ffffff");
    parserFailure.setStackTrace(
        new StackTraceElement[] {
          new StackTraceElement(
              "org.apache.poi.ss.formula.FormulaParser", "nextChar", "FormulaParser.java", 231),
          new StackTraceElement(
              "org.apache.poi.xssf.usermodel.XSSFCell", "setFormula", "XSSFCell.java", 496)
        });

    RuntimeException invalid =
        FormulaExceptions.wrap("Budget", "C7", "[^owe_e`ffffff", parserFailure);

    assertInstanceOf(InvalidFormulaException.class, invalid);
    assertEquals("Budget", ((InvalidFormulaException) invalid).sheetName());
    assertEquals("C7", ((InvalidFormulaException) invalid).address());
    assertEquals("[^owe_e`ffffff", ((InvalidFormulaException) invalid).formula());
    assertSame(parserFailure, invalid.getCause());
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
    assertEquals(Optional.empty(), FormulaExceptions.leadingFunctionName(null));
    assertEquals(Optional.empty(), FormulaExceptions.leadingFunctionName("A1"));
    assertEquals(Optional.empty(), FormulaExceptions.leadingFunctionName(" (A1)"));
    assertEquals(Optional.empty(), FormulaExceptions.leadingFunctionName("1SUM(A1)"));
    assertEquals(
        Optional.of("TEXTAFTER"),
        FormulaExceptions.leadingFunctionName("TEXTAFTER(\"a,b\",\",\")"));
    assertEquals(
        Optional.of("_XLFN.TEXTAFTER"),
        FormulaExceptions.leadingFunctionName("_XLFN.TEXTAFTER(\"a,b\",\",\")"));
  }

  @Test
  void exposesTypedExceptionFieldsAndWorkbookNameFallbacks() {
    MissingExternalWorkbookException missing =
        new MissingExternalWorkbookException(
            "Budget", "D4", "[Rates.xlsx]Sheet1!A1", "Rates.xlsx", "missing", null);
    assertEquals("Budget", missing.sheetName());
    assertEquals("D4", missing.address());
    assertEquals("[Rates.xlsx]Sheet1!A1", missing.formula());
    assertEquals("Rates.xlsx", missing.workbookName());

    UnregisteredUserDefinedFunctionException unregistered =
        new UnregisteredUserDefinedFunctionException(
            "Budget", "E4", "DOUBLE(A1)", "DOUBLE", "missing", null);
    assertEquals("Budget", unregistered.sheetName());
    assertEquals("E4", unregistered.address());
    assertEquals("DOUBLE(A1)", unregistered.formula());
    assertEquals("DOUBLE", unregistered.functionName());

    assertEquals(
        "Rates.xlsx",
        FormulaExceptions.missingExternalWorkbookName(
            new IllegalStateException("no workbook label available"),
            "[Rates.xlsx]Sheet1!A1+[Rates.xlsx]Sheet1!B1+[Other.xlsx]Sheet1!C1"));
    assertNull(
        FormulaExceptions.missingExternalWorkbookName(
            new IllegalStateException("no workbook label available"), null));
    assertEquals(
        List.of("Rates.xlsx", "Other.xlsx"),
        FormulaExceptions.externalWorkbookNames(
            "[Rates.xlsx]Sheet1!A1+[Rates.xlsx]Sheet1!B1+[Other.xlsx]Sheet1!C1"));
  }

  @Test
  void handlesNullMessagesRegisteredFunctionsAndBuiltinCatalogs() throws Exception {
    RuntimeException missingWithoutLabel =
        FormulaExceptions.wrap(
            "Budget",
            "D4",
            "A1",
            new IllegalStateException("outer", new FakeWorkbookNotFoundException(null)));
    assertInstanceOf(MissingExternalWorkbookException.class, missingWithoutLabel);
    assertEquals("Missing external workbook at Budget!D4: A1", missingWithoutLabel.getMessage());

    assertEquals(
        "Rates.xlsx",
        FormulaExceptions.missingExternalWorkbookName(
            new IllegalStateException(null, new FakeWorkbookNotFoundException("ignore me")),
            "[Rates.xlsx]Sheet1!A1"));
    assertEquals(List.of(), FormulaExceptions.externalWorkbookNames(" "));

    ExcelFormulaRuntimeContext registeredContext =
        new ExcelFormulaRuntimeContext(
            Set.of(), ExcelFormulaMissingWorkbookPolicy.ERROR, Set.of("DOUBLE"));
    assertFalse(
        FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
            registeredContext,
            new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException("DOUBLE"),
            "DOUBLE(A1)"));

    assertFalse(
        FormulaExceptions.messageMentionsFunctionName(
            new RuntimeException((String) null), "DOUBLE"));
    assertTrue(
        FormulaExceptions.messageMentionsFunctionName(
            new RuntimeException("function double is unknown"), "DOUBLE"));

    assertTrue(FormulaExceptions.isKnownBuiltinFunction("SUM"));

    ExcelFormulaRuntimeContext emptyContext =
        new ExcelFormulaRuntimeContext(Set.of(), ExcelFormulaMissingWorkbookPolicy.ERROR, Set.of());

    String unsupportedPoiFunction =
        FunctionEval.getNotSupportedFunctionNames().stream().findFirst().orElseThrow();
    assertTrue(FormulaExceptions.isKnownBuiltinFunction(unsupportedPoiFunction));

    String supportedAtpFunction =
        AnalysisToolPak.getSupportedFunctionNames().stream().findFirst().orElseThrow();
    assertTrue(FormulaExceptions.isKnownBuiltinFunction(supportedAtpFunction));

    String unsupportedAtpFunction =
        AnalysisToolPak.getNotSupportedFunctionNames().stream().findFirst().orElseThrow();
    assertTrue(FormulaExceptions.isKnownBuiltinFunction(unsupportedAtpFunction));
    assertFalse(FormulaExceptions.isKnownBuiltinFunction("GRIDGRIND_UNKNOWN_FN"));

    // Cover AnalysisToolPak.getSupportedFunctionNames() as the "first true" OR branch:
    // find an ATP-supported function that is absent from FunctionMetadataRegistry and FunctionEval.
    String atpOnlySupported =
        AnalysisToolPak.getSupportedFunctionNames().stream()
            .filter(n -> FunctionMetadataRegistry.getFunctionByName(n) == null)
            .filter(n -> !FunctionEval.getSupportedFunctionNames().contains(n))
            .filter(n -> !FunctionEval.getNotSupportedFunctionNames().contains(n))
            .findFirst()
            .orElseThrow();
    assertFalse(
        FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
            emptyContext,
            new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException(
                atpOnlySupported),
            atpOnlySupported + "(A1)"));

    // Cover AnalysisToolPak.getNotSupportedFunctionNames() as the "first true" OR branch.
    String atpOnlyNotSupported =
        AnalysisToolPak.getNotSupportedFunctionNames().stream()
            .filter(n -> FunctionMetadataRegistry.getFunctionByName(n) == null)
            .filter(n -> !FunctionEval.getSupportedFunctionNames().contains(n))
            .filter(n -> !FunctionEval.getNotSupportedFunctionNames().contains(n))
            .filter(n -> !AnalysisToolPak.getSupportedFunctionNames().contains(n))
            .findFirst()
            .orElseThrow();
    assertFalse(
        FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
            emptyContext,
            new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException(
                atpOnlyNotSupported),
            atpOnlyNotSupported + "(A1)"));
  }

  /** Test-only stand-in used to exercise workbook-missing classification. */
  private static final class FakeWorkbookNotFoundException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    private FakeWorkbookNotFoundException(String message) {
      super(message);
    }
  }
}
