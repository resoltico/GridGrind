package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.plan.ExecutionJournalInput;
import dev.erst.gridgrind.core.plan.ExecutionModeInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/**
 * Verifies that Java authoring preserves canonical execution policy while changing journal level.
 */
class GridGrindPlanExecutionPolicyTest {
  @Test
  void journalPreservesImportedExecutionModeAndCalculation() {
    ExecutionModeInput mode = ExecutionModeInput.eventRead();
    dev.erst.gridgrind.core.model.formula.CalculationPolicyInput calculation =
        new dev.erst.gridgrind.core.model.formula.CalculationPolicyInput(
            new dev.erst.gridgrind.core.model.formula.CalculationStrategyInput.DoNotCalculate(),
            true);
    WorkbookPlan canonical =
        new WorkbookPlan(
            dev.erst.gridgrind.core.plan.GridGrindProtocolVersion.current(),
            Optional.of(new PlanId("canonical-plan")),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicyInput(
                mode,
                new ExecutionJournalInput(ExecutionJournalLevel.NORMAL),
                calculation,
                dev.erst.gridgrind.core.plan.AssertionModeInput.defaults()),
            FormulaEnvironmentInput.empty(),
            List.of());

    WorkbookPlan journaled =
        GridGrindPlan.from(canonical).journal(ExecutionJournalLevel.VERBOSE).toPlan();

    assertEquals(mode, journaled.execution().mode());
    assertEquals(calculation, journaled.execution().calculation());
    assertEquals(ExecutionJournalLevel.VERBOSE, journaled.execution().journal().level());
  }
}
