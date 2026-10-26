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

import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.value.CellScalarValue;
import dev.erst.gridgrind.core.plan.AssertionModeSpec;
import dev.erst.gridgrind.core.plan.AssertionOutcome;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
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
                    ExecutionPolicySpec.assertionMode(AssertionModeSpec.COLLECT),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName(new SheetName("Budget")),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress(
                                new SheetName("Budget"), new CellAddress("A1")),
                            new CellMutationAction.SetCell(textCell("Owner")))),
                    assertions(
                        assertThat(
                            "assert-pass",
                            new CellSelector.ByAddress(
                                new SheetName("Budget"), new CellAddress("A1")),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Owner"))),
                        assertThat(
                            "assert-first",
                            new CellSelector.ByAddress(
                                new SheetName("Budget"), new CellAddress("A1")),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Wrong first"))),
                        assertThat(
                            "assert-second",
                            new CellSelector.ByAddress(
                                new SheetName("Budget"), new CellAddress("A1")),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Wrong second")))),
                    inspections(
                        inspect(
                            "cells-after-assertions",
                            new CellSelector.ByAddresses(
                                new SheetName("Budget"), List.of(new CellAddress("A1"))),
                            allFacetCellsQuery())))));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertEquals("assert-first", failedAssertion(failure).stepId().value());
    assertEquals(
        List.of(AssertionOutcome.PASSED, AssertionOutcome.FAILED, AssertionOutcome.FAILED),
        failure.assertions().stream().map(AssertionResult::outcome).toList());
    assertEquals(
        List.of("cells-after-assertions"),
        failure.inspections().stream().map(InspectionResult::stepId).toList().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
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
                    new ExecutionPolicySpec(
                        ExecutionModeSpec.streamingWrite(),
                        ExecutionJournalSpec.defaults(),
                        CalculationPolicySpec.defaults(),
                        AssertionModeSpec.COLLECT),
                    null,
                    mutations(
                        mutate(
                            new SheetSelector.ByName(new SheetName("Stream")),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName(new SheetName("Stream")),
                            new CellMutationAction.AppendRow(
                                new dev.erst.gridgrind.core.model.value.CellRowSpec.Typed(
                                    List.of(textCell("Owner"), textCell("Ada")))))),
                    assertions(
                        assertThat(
                            "stream-pass",
                            new CellSelector.ByAddress(
                                new SheetName("Stream"), new CellAddress("A1")),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Owner"))),
                        assertThat(
                            "stream-fail",
                            new CellSelector.ByAddress(
                                new SheetName("Stream"), new CellAddress("A1")),
                            new CellAssertion.CellValue(new CellScalarValue.Text("Wrong")))),
                    inspections(
                        inspect(
                            "stream-after-failure",
                            new CellSelector.ByAddress(
                                new SheetName("Stream"), new CellAddress("A1")),
                            allFacetCellsQuery())))));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertEquals("stream-fail", failedAssertion(failure).stepId().value());
    assertEquals(
        List.of(AssertionOutcome.PASSED, AssertionOutcome.FAILED),
        failure.assertions().stream().map(AssertionResult::outcome).toList());
    assertEquals(
        List.of("stream-after-failure"),
        failure.inspections().stream().map(InspectionResult::stepId).toList().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
  }

  private static dev.erst.gridgrind.core.plan.AssertionFailure failedAssertion(
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
