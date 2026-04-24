package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
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
    assertTrue(GridGrindContractText.stepKindSummary().contains("stable caller-defined stepId"));
    assertTrue(
        GridGrindContractText.workbookFindingsDiscoverySummary()
            .contains("ANALYZE_WORKBOOK_FINDINGS"));
    assertEquals(
        "execution.mode.readMode=EVENT_READ requires execution.calculation.strategy="
            + "DO_NOT_CALCULATE and markRecalculateOnOpen=false",
        GridGrindExecutionModeMetadata.eventRead().calculationFailureMessage());
    assertEquals(
        "execution.mode.writeMode=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW only;"
            + " unsupported mutation action type: SET_CELL",
        GridGrindExecutionModeMetadata.streamingWrite().unsupportedActionMessage("SET_CELL"));
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
  void helperBranchesStayCovered() {
    assertEquals("only", GridGrindContractText.humanJoin(List.of("only")));
    assertEquals("left and right", GridGrindContractText.humanJoin(List.of("left", "right")));
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        GridGrindContractText.typeNamesByClass(MutationAction.class)
            .get(MutationAction.ClearWorkbookProtection.class));
    assertEquals(
        "left, middle, and right",
        GridGrindContractText.humanJoin(List.of("left", "middle", "right")));
    assertEquals(
        "values must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindContractText.humanJoin(List.of("", " ")))
            .getMessage());
  }

  @Test
  void executionModeMetadataRejectsEmptyAllowedLists() {
    assertEquals(
        "allowedQueries must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExecutionModeMetadata.EventReadMode(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        List.of(),
                        CalculationStrategyInput.DoNotCalculate.class,
                        false))
            .getMessage());
    assertEquals(
        "allowedActions must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindExecutionModeMetadata.StreamingWriteMode(
                        ExecutionModeInput.WriteMode.STREAMING_WRITE,
                        WorkbookPlan.WorkbookSource.New.class,
                        List.of(),
                        CalculationStrategyInput.DoNotCalculate.class,
                        true))
            .getMessage());
  }
}
