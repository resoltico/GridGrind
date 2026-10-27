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
    var readyFirst = new PreparedMutationStep(first, action);
    var readySecond = new PreparedMutationStep(second, action);
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
        () -> new PreparedMutationStep(first, new WorkbookMutationAction.DeleteSheet()));
  }
}
