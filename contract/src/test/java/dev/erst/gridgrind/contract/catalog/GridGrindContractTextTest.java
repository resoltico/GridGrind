package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Tests for core-owned contract wording shared by downstream discovery surfaces. */
class GridGrindContractTextTest {
  @Test
  void resolvesCanonicalTypeNamesAndPhrases() {
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        GridGrindContractText.mutationActionTypeName(MutationAction.ClearWorkbookProtection.class));
    assertEquals(
        "GET_SHEET_SUMMARY",
        GridGrindContractText.inspectionQueryTypeName(InspectionQuery.GetSheetSummary.class));
    assertEquals(
        Set.of(MutationAction.EnsureSheet.class, MutationAction.AppendRow.class),
        GridGrindContractText.streamingWriteMutationActionClasses());
    assertEquals(
        Set.of(InspectionQuery.GetWorkbookSummary.class, InspectionQuery.GetSheetSummary.class),
        GridGrindContractText.eventReadInspectionQueryClasses());
    assertTrue(GridGrindContractText.executionModeInputSummary().contains("markRecalculateOnOpen"));
    assertTrue(GridGrindContractText.sheetLayoutReadSummary().contains("presentation"));
    assertTrue(
        GridGrindContractText.workbookFindingsDiscoverySummary()
            .contains("ANALYZE_WORKBOOK_FINDINGS"));
  }

  @Test
  void rejectsUnknownSubtypeClasses() {
    assertThrows(
        NullPointerException.class, () -> GridGrindContractText.mutationActionTypeName(null));
    assertThrows(
        NullPointerException.class, () -> GridGrindContractText.inspectionQueryTypeName(null));

    @SuppressWarnings("unchecked")
    Class<? extends MutationAction> unsupportedActionClass =
        (Class<? extends MutationAction>) (Class<?>) GridGrindContractTextTest.class;
    @SuppressWarnings("unchecked")
    Class<? extends InspectionQuery> unsupportedQueryClass =
        (Class<? extends InspectionQuery>) (Class<?>) GridGrindContractTextTest.class;

    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindContractText.mutationActionTypeName(unsupportedActionClass))
            .getMessage()
            .contains("MutationAction"));
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindContractText.inspectionQueryTypeName(unsupportedQueryClass))
            .getMessage()
            .contains("InspectionQuery"));
  }

  @Test
  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  void privateHelpersCoverDefensiveBranches() throws Exception {
    Method humanJoin = GridGrindContractText.class.getDeclaredMethod("humanJoin", List.class);
    humanJoin.setAccessible(true);
    assertEquals("only", invokeString(humanJoin, List.of("only")));
    assertEquals("left and right", invokeString(humanJoin, List.of("left", "right")));
    assertEquals(
        "left, middle, and right", invokeString(humanJoin, List.of("left", "middle", "right")));
    assertEquals(
        "values must not be empty",
        assertThrows(
                IllegalArgumentException.class, () -> invokeString(humanJoin, List.of("", " ")))
            .getMessage());
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
