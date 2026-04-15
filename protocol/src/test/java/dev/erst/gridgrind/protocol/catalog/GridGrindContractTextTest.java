package dev.erst.gridgrind.protocol.catalog;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Tests for core-owned public contract wording shared by downstream discovery surfaces. */
class GridGrindContractTextTest {
  @Test
  void resolvesProtocolTypeNamesFromTheCanonicalSubtypeRegistry() {
    assertEquals(
        "FORCE_FORMULA_RECALCULATION_ON_OPEN",
        GridGrindContractText.operationTypeName(
            WorkbookOperation.ForceFormulaRecalculationOnOpen.class));
    assertEquals(
        "GET_SHEET_SUMMARY",
        GridGrindContractText.readTypeName(WorkbookReadOperation.GetSheetSummary.class));
  }

  @Test
  void exposesStableHumanReadableContractPhrases() {
    assertEquals(
        Set.of(
            WorkbookOperation.EnsureSheet.class,
            WorkbookOperation.AppendRow.class,
            WorkbookOperation.ForceFormulaRecalculationOnOpen.class),
        GridGrindContractText.streamingWriteOperationClasses());
    assertEquals(
        Set.of(
            WorkbookReadOperation.GetWorkbookSummary.class,
            WorkbookReadOperation.GetSheetSummary.class),
        GridGrindContractText.eventReadOperationClasses());
    assertEquals(
        "ENSURE_SHEET, APPEND_ROW, and FORCE_FORMULA_RECALCULATION_ON_OPEN",
        GridGrindContractText.streamingWriteOperationTypePhrase());
    assertEquals(
        "GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY",
        GridGrindContractText.eventReadReadTypePhrase());
    assertEquals(
        "formula health, data-validation health, conditional-formatting health,"
            + " autofilter health, table health, pivot-table health, hyperlink health,"
            + " and named-range health",
        GridGrindContractText.workbookAnalysisFamilyPhrase());
    assertTrue(GridGrindContractText.formulaAuthoringLimitSummary().contains("LAMBDA/LET"));
    assertTrue(GridGrindContractText.loadedFormulaSupportSummary().contains("UNSUPPORTED_FORMULA"));
    assertTrue(GridGrindContractText.sheetLayoutReadSummary().contains("presentation"));
    assertTrue(
        GridGrindContractText.formulaSurfaceReadSummary()
            .contains("analysis.totalFormulaCellCount"));
    assertTrue(
        GridGrindContractText.namedRangeSurfaceReadSummary()
            .contains("analysis.workbookScopedCount"));
    assertTrue(
        GridGrindContractText.formulaHealthReadSummary()
            .contains("analysis.checkedFormulaCellCount"));
    assertTrue(
        GridGrindContractText.namedRangeHealthReadSummary()
            .contains("analysis.checkedNamedRangeCount"));
    assertTrue(GridGrindContractText.workbookFindingsReadSummary().contains("analysis.findings"));
    assertTrue(
        GridGrindContractText.executionModeInputSummary()
            .contains("FORCE_FORMULA_RECALCULATION_ON_OPEN"));
    assertTrue(
        GridGrindContractText.workbookFindingsDiscoverySummary()
            .contains("ANALYZE_WORKBOOK_FINDINGS"));
  }

  @Test
  void rejectsUnknownOrNullProtocolSubtypeClasses() {
    assertThrows(NullPointerException.class, () -> GridGrindContractText.operationTypeName(null));
    assertThrows(NullPointerException.class, () -> GridGrindContractText.readTypeName(null));

    @SuppressWarnings("unchecked")
    Class<? extends WorkbookOperation> unsupportedOperationClass =
        (Class<? extends WorkbookOperation>) (Class<?>) GridGrindContractTextTest.class;
    @SuppressWarnings("unchecked")
    Class<? extends WorkbookReadOperation> unsupportedReadClass =
        (Class<? extends WorkbookReadOperation>) (Class<?>) GridGrindContractTextTest.class;

    IllegalArgumentException operationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindContractText.operationTypeName(unsupportedOperationClass));
    assertTrue(operationFailure.getMessage().contains("WorkbookOperation"));

    IllegalArgumentException readFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindContractText.readTypeName(unsupportedReadClass));
    assertTrue(readFailure.getMessage().contains("WorkbookReadOperation"));
  }

  @Test
  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  void privateHelpersCoverTheDefensiveBranches() throws Exception {
    Method humanJoin = GridGrindContractText.class.getDeclaredMethod("humanJoin", List.class);
    humanJoin.setAccessible(true);
    assertEquals("only", invokeString(humanJoin, List.of("only")));
    assertEquals("left and right", invokeString(humanJoin, List.of("left", "right")));
    IllegalArgumentException emptyJoinFailure =
        assertThrows(
            IllegalArgumentException.class, () -> invokeString(humanJoin, List.of("", " ")));
    assertEquals("values must not be empty", emptyJoinFailure.getMessage());

    Method typeNamesByClass =
        GridGrindContractText.class.getDeclaredMethod("typeNamesByClass", Class.class);
    typeNamesByClass.setAccessible(true);
    IllegalArgumentException missingAnnotationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> invoke(typeNamesByClass, GridGrindContractTextTest.class));
    assertTrue(missingAnnotationFailure.getMessage().contains("@JsonSubTypes"));
  }

  private static String invokeString(Method method, Object argument) {
    return (String) invoke(method, argument);
  }

  @SuppressWarnings("PMD.PreserveStackTrace")
  private static Object invoke(Method method, Object argument) {
    try {
      return method.invoke(null, argument);
    } catch (IllegalAccessException exception) {
      throw new AssertionError(exception);
    } catch (InvocationTargetException exception) {
      Throwable cause = exception.getCause();
      if (cause instanceof RuntimeException runtimeException) {
        throw runtimeException;
      }
      if (cause instanceof Error error) {
        throw error;
      }
      throw new AssertionError(cause);
    }
  }
}
