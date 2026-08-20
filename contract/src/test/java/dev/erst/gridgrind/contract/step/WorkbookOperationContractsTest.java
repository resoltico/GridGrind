package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.assertion.AnalysisAssertion;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Boundary coverage for the canonical operation-target contract registry. */
class WorkbookOperationContractsTest {
  @Test
  void rejectsDynamicContractsWhenAStaticSelectorListWasRequested() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookOperationContracts.staticTargetSelectorsFor(
                    AnalysisAssertion.AnalysisMaxSeverity.class));

    assertEquals(
        "Operation type "
            + AnalysisAssertion.AnalysisMaxSeverity.class.getName()
            + " derives target selectors dynamically",
        failure.getMessage());
  }

  @Test
  void rejectsTypesOutsideTheSealedWorkbookOperationFamilies() {
    IllegalArgumentException nonRecordFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> WorkbookOperationContracts.targetSurfaceFor(String.class));
    assertEquals(
        "Operation type must be a record: class java.lang.String", nonRecordFailure.getMessage());

    IllegalArgumentException unrelatedRecordFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> WorkbookOperationContracts.targetSurfaceFor(UnrelatedRecord.class));
    assertEquals(
        "Operation type must be a record: class " + UnrelatedRecord.class.getName(),
        unrelatedRecordFailure.getMessage());
  }

  @Test
  void rejectsEmptyOrUndeclaredTargetContractFacts() {
    assertEquals(
        "selectorTypes must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new StaticWorkbookOperationTargetContract(List.of()))
            .getMessage());

    WorkbookOperationTargetSelectorDerivation derivation = ignored -> selectorTypes(Selector.class);
    assertEquals(
        "rule must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DerivedWorkbookOperationTargetContract(null, derivation))
            .getMessage());
    assertEquals(
        "rule must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DerivedWorkbookOperationTargetContract(" ", derivation))
            .getMessage());
  }

  @Test
  void namesEveryAcceptedTargetFamilyInAStaticMismatch() {
    assertEquals(
        "EXPECT_CELL_VALUE requires target type CELL_BY_ADDRESS, CELL_BY_ADDRESSES or "
            + "TABLE_CELL_BY_COLUMN_NAME but got WORKBOOK_CURRENT",
        WorkbookOperationContracts.targetViolation(
                new CellAssertion.CellValue(new CellScalarValue.Text("Owner")),
                new WorkbookSelector.Current())
            .orElseThrow());
  }

  @Test
  void ownsOperationModeCompatibilityAlongsideTargetCompatibility() {
    assertEquals(
        "execution.mode.type=EVENT_READ supports inspection steps only; unsupported step kind: MUTATION",
        WorkbookOperationContracts.executionModeViolation(
                new CellMutationAction.SetCell(new CellInput.NumberValue(1.0)),
                new ExecutionModeInput.EventRead())
            .orElseThrow());
  }

  @SuppressWarnings("unchecked")
  private static Class<? extends Selector>[] selectorTypes(Class<? extends Selector> selectorType) {
    return new Class[] {selectorType};
  }

  private record UnrelatedRecord() {}
}
