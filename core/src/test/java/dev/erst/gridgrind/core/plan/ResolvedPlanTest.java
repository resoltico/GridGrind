package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Prepared execution retains ordered identities while owning its materialized step list. */
class ResolvedPlanTest {
  @Test
  void ownsPreparedStepsAndRejectsOrderOrActionIdentityMismatch() {
    var action = new WorkbookMutationAction.EnsureSheet();
    var first =
        new MutationStep(
            new StepId("first"), new SheetSelector.ByName(new SheetName("Budget")), action);
    var second =
        new MutationStep(
            new StepId("second"), new SheetSelector.ByName(new SheetName("Budget")), action);
    var request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(first, second));
    var readyFirst = (PreparedMutationStep) PreparedStep.prepare(first);
    var readySecond = (PreparedMutationStep) PreparedStep.prepare(second);
    var input = new ArrayList<PreparedStep>(List.of(readyFirst, readySecond));
    var plan = new ResolvedPlan(request, input);
    input.clear();
    assertEquals(2, plan.steps().size());
    assertSame(action, ((PreparedMutationStep) plan.steps().getFirst()).action());
    assertSame(first, plan.steps().getFirst().authoredStep());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedPlan(request, List.of(readySecond, readyFirst)));
    assertThrows(
        IllegalArgumentException.class, () -> new ResolvedPlan(request, List.of(readyFirst)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                first,
                new MutationStep(
                    first.stepId(), first.target(), new WorkbookMutationAction.DeleteSheet())));
  }

  @Test
  void rejectsUnrelatedPreparedStepsSharingAnAuthoredIdentity() {
    var id = new StepId("same");
    var target =
        new CellSelector.ByAddress(
            new SheetName("Budget"), new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"));
    var original =
        new AssertionStep(
            id,
            target,
            new CellAssertion.Value(
                new dev.erst.gridgrind.core.model.value.CellValue.Text("Original")));
    var changed =
        new AssertionStep(
            id,
            target,
            new CellAssertion.Value(
                new dev.erst.gridgrind.core.model.value.CellValue.Text("Unrelated")));
    var request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(original));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedPlan(request, List.of(PreparedStep.prepare(changed))));
  }
}
