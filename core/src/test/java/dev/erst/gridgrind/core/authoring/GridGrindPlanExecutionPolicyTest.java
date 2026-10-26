package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
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
    ExecutionModeSpec mode = ExecutionModeSpec.eventRead();
    dev.erst.gridgrind.core.model.formula.CalculationPolicySpec calculation =
        new dev.erst.gridgrind.core.model.formula.CalculationPolicySpec(
            new dev.erst.gridgrind.core.model.formula.CalculationStrategySpec.DoNotCalculate(),
            true);
    WorkbookPlan canonical =
        new WorkbookPlan(
            dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT,
            Optional.of(new PlanId("canonical-plan")),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicySpec(
                mode,
                new ExecutionJournalSpec(ExecutionJournalLevel.NORMAL),
                calculation,
                dev.erst.gridgrind.core.plan.AssertionModeSpec.defaults()),
            FormulaEnvironmentSpec.empty(),
            List.of());

    WorkbookPlan journaled =
        GridGrindPlan.from(canonical).journal(ExecutionJournalLevel.VERBOSE).toPlan();

    assertEquals(mode, journaled.execution().mode());
    assertEquals(calculation, journaled.execution().calculation());
    assertEquals(ExecutionJournalLevel.VERBOSE, journaled.execution().journal().level());
  }
}
