package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies a later formula failure identifies the prior step that authored the affected cell. */
class DefaultGridGrindRequestExecutorFormulaOriginTest {
  @TempDir Path root;

  @Test
  void attributesFormulaFailuresToTheirEarlierAuthoringStep() throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure-ops"),
                    new SheetSelector.ByName(new SheetName("Ops")),
                    new WorkbookMutationAction.EnsureSheet()),
                new MutationStep(
                    new StepId("author-formula"),
                    new CellSelector.ByAddress(new SheetName("Ops"), "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(TextSourceInput.inline("1+1")))),
                new MutationStep(
                    new StepId("reject-formula"),
                    new CellSelector.ByAddress(new SheetName("Ops"), "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(TextSourceInput.inline("SUM("))))));

    WorkbookResult result =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request,
                new ExecutionInputBindings(
                    root,
                    Files.createDirectory(root.resolve("temp")),
                    ExecutionGrantTestSupport.noPublication(
                        List.of("ENSURE_SHEET", "SET_CELL"), List.of())),
                ExecutionProgressSink.NOOP);

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, result);
    ProblemContext.ExecuteStep context =
        assertInstanceOf(ProblemContext.ExecuteStep.class, failure.problem().context());
    assertEquals("author-formula", context.stepId().value());
    assertEquals("reject-formula", context.surfacedAtStep().orElseThrow().stepId().value());
  }
}
