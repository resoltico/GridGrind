package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.assertThat;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.assertions;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.inspect;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.inspections;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.textCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.AssertionOutcome;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Integration coverage for terminal-phase assertion collection. */
class AssertionCollectionWorkflowTest extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void collectsTerminalAssertionFailuresBeforeReturningTheCanonicalFirstFailure() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionPolicyInput.assertionMode(AssertionModeInput.COLLECT),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellMutationAction.SetCell(textCell("Owner")))),
                    assertions(
                        assertThat(
                            "assert-pass",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Owner"))),
                        assertThat(
                            "assert-first",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Wrong first"))),
                        assertThat(
                            "assert-second",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Wrong second")))),
                    inspections(
                        inspect(
                            "cells-after-assertions",
                            new CellSelector.ByAddresses("Budget", List.of("A1")),
                            allFacetCellsQuery())))));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertEquals("assert-first", failedAssertion(failure).stepId());
    assertEquals(
        List.of(AssertionOutcome.PASSED, AssertionOutcome.FAILED, AssertionOutcome.FAILED),
        failure.assertions().stream().map(AssertionResult::outcome).toList());
    assertEquals(
        List.of("cells-after-assertions"),
        failure.inspections().stream().map(InspectionResult::stepId).toList());
    assertInstanceOf(
        WorkbookResultPersistence.PersistenceOutcome.NotSaved.class, failure.persistence());
  }

  @Test
  void collectsTerminalAssertionFailuresInStreamingWriteMode() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    new ExecutionPolicyInput(
                        ExecutionModeInput.streamingWrite(),
                        ExecutionJournalInput.defaults(),
                        CalculationPolicyInput.defaults(),
                        AssertionModeInput.COLLECT),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Stream"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Stream"),
                            new CellMutationAction.AppendRow(
                                new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                    List.of(textCell("Owner"), textCell("Ada")))))),
                    assertions(
                        assertThat(
                            "stream-pass",
                            new CellSelector.ByAddress("Stream", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Owner"))),
                        assertThat(
                            "stream-fail",
                            new CellSelector.ByAddress("Stream", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Wrong")))),
                    inspections(
                        inspect(
                            "stream-after-failure",
                            new CellSelector.ByAddress("Stream", "A1"),
                            allFacetCellsQuery())))));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertEquals("stream-fail", failedAssertion(failure).stepId());
    assertEquals(
        List.of(AssertionOutcome.PASSED, AssertionOutcome.FAILED),
        failure.assertions().stream().map(AssertionResult::outcome).toList());
    assertEquals(
        List.of("stream-after-failure"),
        failure.inspections().stream().map(InspectionResult::stepId).toList());
  }

  private static dev.erst.gridgrind.contract.assertion.AssertionFailure failedAssertion(
      WorkbookResult.Failure failure) {
    return assertInstanceOf(
            AssertionResult.Failed.class,
            failure.assertions().stream()
                .filter(result -> result.outcome() == AssertionOutcome.FAILED)
                .findFirst()
                .orElseThrow())
        .failure();
  }
}
