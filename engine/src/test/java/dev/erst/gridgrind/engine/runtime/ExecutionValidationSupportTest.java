package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies execution validation preserves every independently provable authored violation. */
class ExecutionValidationSupportTest {
  @Test
  void retainsEquivalentStreamingWriteFailuresInAuthoredStepOrder() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(new ExecutionModeInput.StreamingWrite()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "ensure-sheet",
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.EnsureSheet()),
                setCell("set-first", "A1"),
                setCell("set-second", "A2")));

    List<String> messages =
        new ExecutionValidationSupport()
            .validateRequest(request).stream().map(problem -> problem.message()).toList();
    String unsupportedSetCell =
        GridGrindExecutionModeMetadata.streamingWrite().unsupportedActionMessage("SET_CELL");

    assertEquals(List.of(unsupportedSetCell, unsupportedSetCell), messages);
  }

  private static MutationStep setCell(String stepId, String address) {
    return new MutationStep(
        stepId,
        new CellSelector.ByAddress("Budget", address),
        new CellMutationAction.SetCell(new CellInput.NumberValue(1.0d)));
  }
}
