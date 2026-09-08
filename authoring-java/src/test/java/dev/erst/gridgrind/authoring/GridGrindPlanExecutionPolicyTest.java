package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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
    dev.erst.gridgrind.contract.dto.CalculationPolicyInput calculation =
        new dev.erst.gridgrind.contract.dto.CalculationPolicyInput(
            new dev.erst.gridgrind.contract.dto.CalculationStrategyInput.DoNotCalculate(), true);
    WorkbookPlan canonical =
        new WorkbookPlan(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            Optional.of("canonical-plan"),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicyInput(
                mode,
                new ExecutionJournalInput(ExecutionJournalLevel.NORMAL),
                calculation,
                dev.erst.gridgrind.contract.dto.AssertionModeInput.defaults()),
            FormulaEnvironmentInput.empty(),
            List.of());

    WorkbookPlan journaled =
        GridGrindPlan.from(canonical).journal(ExecutionJournalLevel.VERBOSE).toPlan();

    assertEquals(mode, journaled.execution().mode());
    assertEquals(calculation, journaled.execution().calculation());
    assertEquals(ExecutionJournalLevel.VERBOSE, journaled.execution().journal().level());
  }
}
