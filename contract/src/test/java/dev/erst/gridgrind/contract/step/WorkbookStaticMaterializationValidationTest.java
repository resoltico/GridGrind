package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies static rejection of row operations that would create excessive worksheet records. */
class WorkbookStaticMaterializationValidationTest {
  @Test
  void materializationValidatorRemainsInternalOnly() throws ReflectiveOperationException {
    java.lang.reflect.Constructor<WorkbookStaticMaterializationValidation> constructor =
        WorkbookStaticMaterializationValidation.class.getDeclaredConstructor();

    assertFalse(constructor.canAccess(null));
    assertTrue(constructor.trySetAccessible());
    assertEquals(
        WorkbookStaticMaterializationValidation.class, constructor.newInstance().getClass());
  }

  @Test
  void rejectsEveryRowMaterializingOperationBeyondTheCumulativeWorkBudget() {
    assertAll(
        List.of(
                new WorkbookMutationAction.SetRowHeight(15.0d),
                new WorkbookMutationAction.SetRowVisibility(true),
                WorkbookMutationAction.GroupRows.expanded(),
                new WorkbookMutationAction.UngroupRows())
            .stream()
            .<org.junit.jupiter.api.function.Executable>map(
                action -> () -> assertRowWorkRejection(action))
            .toList());
  }

  @Test
  void acceptsTheExactRowWorkBudget() {
    assertEquals(
        List.of(),
        WorkbookStaticRequestContract.validate(
            staticRequest(
                new MutationStep(
                    "row-work",
                    new RowBandSelector.Span("Ops", 10, 250_009),
                    new WorkbookMutationAction.SetRowHeight(15.0d)))));
  }

  @Test
  void doesNotChargeCellActionsAgainstRowBandMaterializationWork() {
    List<WorkbookStaticViolation> violations =
        WorkbookStaticRequestContract.validate(
            staticRequest(
                new MutationStep(
                    "row-band-cell-action",
                    new RowBandSelector.Span("Ops", 10, 250_010),
                    new CellMutationAction.ClearRange())));

    assertFalse(
        violations.stream().anyMatch(violation -> violation.message().contains("materialized")));
  }

  private static void assertRowWorkRejection(WorkbookMutationAction action) {
    WorkbookStaticViolation violation =
        WorkbookStaticRequestContract.validate(
                staticRequest(
                    new MutationStep(
                        "row-work", new RowBandSelector.Span("Ops", 10, 250_010), action)))
            .getFirst();

    assertEquals("steps[0].target.lastRowIndex", violation.jsonPath());
    assertEquals(
        "materialized worksheet work must not exceed 250000 items per plan; this step raises the total to 250001",
        violation.message());
  }

  private static WorkbookStaticRequest staticRequest(MutationStep step) {
    return new WorkbookStaticRequest(
        Optional.of(new WorkbookPlan.WorkbookSource.New()),
        Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
        Optional.of(ExecutionPolicyInput.defaults()),
        List.of(new WorkbookStaticStep(0, Optional.of(step))));
  }
}
