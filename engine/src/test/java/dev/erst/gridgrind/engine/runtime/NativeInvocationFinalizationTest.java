package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.*;
import dev.erst.gridgrind.core.model.value.*;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.*;
import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import dev.erst.gridgrind.core.vocabulary.*;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** The native owner finalizes before terminal telemetry and preserves primary failures. */
@SuppressWarnings("PMD.CloseResource")
// Executor ownership must close both the real context and injected teardown before terminal
// callbacks.
// A surrounding TWR would conceal that contract from these finalization assertions.
class NativeInvocationFinalizationTest {
  @TempDir Path root;

  @Test
  void cleanupFailurePrecedesTheOnlyTerminalEventAndPreservesCompletedFacts() {
    var order = new ArrayList<String>();
    var real = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD);
    var context =
        new NativeInvocationContext(
            List.of(
                real::close,
                () -> {
                  order.add("teardown");
                  throw new IllegalStateException("owned cleanup failure");
                }));
    var plan = plan(false);
    var events = new ArrayList<ExecutionProgressEvent>();
    var result =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor(context)
                .execute(
                    plan,
                    ExecutionInputBindingsFixtureSupport.bindings(root, plan),
                    event -> {
                      events.add(event);
                      if (event.category() == ExecutionProgressEvent.Category.PLAN
                          && event.status() != ExecutionProgressEvent.Status.STARTED)
                        order.add("terminal");
                    }));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, result.problem().code());
    assertEquals(List.of("teardown", "terminal"), order);
    assertEquals(ExecutionProgressEvent.Status.FAILED, events.getLast().status());
    assertEquals(
        2,
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, result.journal().outcome())
            .completedStepCount());
    assertEquals(
        1,
        events.stream()
            .filter(e ->
                e.category() == ExecutionProgressEvent.Category.PLAN
                    && e.status() != ExecutionProgressEvent.Status.STARTED)
            .count());
  }

  @Test
  void existingAssertionFailureRemainsPrimaryWhenNativeTeardownFails() {
    var real = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD);
    var context =
        new NativeInvocationContext(
            List.of(
                real::close,
                () -> {
                  throw new IllegalStateException("owned cleanup failure");
                }));
    var plan = plan(true);
    var result =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(context), plan, root));
    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, result.problem().code());
    assertEquals(new StepId("assert"), result.assertions().getFirst().stepId());
    assertTrue(
        result.problem().causes().stream()
            .anyMatch(c -> c.code() == GridGrindProblemCode.INTERNAL_ERROR));
    assertEquals(
        new StepId("assert"),
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, result.journal().outcome())
            .failedStep()
            .orElseThrow()
            .failedStepId());
  }

  @Test
  void primaryErrorIdentityAndInterruptionSurviveSecondaryTeardownFailure() {
    var real = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD);
    var secondary = new IllegalStateException("owned cleanup failure");
    var context =
        new NativeInvocationContext(
            List.of(
                real::close,
                () -> {
                  throw secondary;
                }));
    var primary = new AssertionError("owned primary failure");
    var plan = plan(false);
    var error =
        assertThrows(
            AssertionError.class,
            () ->
                new DefaultGridGrindRequestExecutor(context)
                    .execute(
                        plan,
                        ExecutionInputBindingsFixtureSupport.bindings(root, plan),
                        event -> {
                          if (event.stepId().map(id -> "write".equals(id.value())).orElse(false)
                              && event.status() == ExecutionProgressEvent.Status.SUCCEEDED) {
                            Thread.currentThread().interrupt();
                            throw primary;
                          }
                        }));
    try {
      assertSame(primary, error);
      assertEquals(1, error.getSuppressed().length);
      assertSame(secondary, error.getSuppressed()[0]);
      assertTrue(Thread.currentThread().isInterrupted());
    } finally {
      Thread.interrupted();
    }
  }

  @Test
  void cleanupErrorKeepsIdentityAndReceivesRuntimeFailuresInOrder() {
    var first = new IllegalStateException("first cleanup failure");
    var fatal = new AssertionError("fatal cleanup failure");
    var last = new IllegalStateException("last cleanup failure");
    var context =
        new NativeInvocationContext(
            List.of(
                () -> {
                  throw first;
                },
                () -> {
                  throw fatal;
                },
                () -> {
                  throw last;
                }));
    var observed = assertThrows(AssertionError.class, context::close);
    assertSame(fatal, observed);
    assertEquals(2, observed.getSuppressed().length);
    assertSame(first, observed.getSuppressed()[0]);
    assertSame(last, observed.getSuppressed()[1]);
  }

  private static WorkbookPlan plan(boolean assertion) {
    var sheet = new SheetName("Data");
    var steps = new ArrayList<WorkbookStep>();
    steps.add(
        new MutationStep(
            new StepId("sheet"),
            new SheetSelector.ByName(sheet),
            new WorkbookMutationAction.EnsureSheet()));
    steps.add(
        new MutationStep(
            new StepId("write"),
            new CellSelector.ByAddress(sheet, new CellAddress("A1")),
            new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1))));
    if (assertion)
      steps.add(
          new AssertionStep(
              new StepId("assert"),
              new CellSelector.ByAddress(sheet, new CellAddress("A1")),
              new CellAssertion.Value(new CellValue.NumberValue(2))));
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        new ExecutionPolicySpec(
            ExecutionModeSpec.defaults(),
            new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
            CalculationPolicySpec.defaults(),
            AssertionModeSpec.FAIL_FAST),
        FormulaEnvironmentSpec.empty(),
        steps);
  }
}
