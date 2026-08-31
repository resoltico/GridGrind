package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
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
                    "ensure-ops",
                    new SheetSelector.ByName("Ops"),
                    new WorkbookMutationAction.EnsureSheet()),
                new MutationStep(
                    "author-formula",
                    new CellSelector.ByAddress("Ops", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(TextSourceInput.inline("1+1")))),
                new MutationStep(
                    "reject-formula",
                    new CellSelector.ByAddress("Ops", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Formula(TextSourceInput.inline("SUM("))))));

    WorkbookResult result =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request,
                new ExecutionInputBindings(root, Files.createDirectory(root.resolve("temp"))),
                ExecutionProgressSink.NOOP);

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, result);
    ProblemContext.ExecuteStep context =
        assertInstanceOf(ProblemContext.ExecuteStep.class, failure.problem().context());
    assertEquals("author-formula", context.stepId());
    assertEquals("reject-formula", context.surfacedAtStep().orElseThrow().stepId());
  }
}
