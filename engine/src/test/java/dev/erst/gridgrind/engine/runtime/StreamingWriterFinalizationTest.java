package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.AssertionModeSpec;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempScope;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

/** A native writer-close failure is secondary to an established execution refusal. */
class StreamingWriterFinalizationTest {
  @TempDir Path root;

  @ParameterizedTest
  @EnumSource(
      value = CloseFailure.class,
      names = {"IO", "UNCHECKED"})
  void retainsMutationFailureWhenTheNativeWriterAlsoFailsToClose(CloseFailure closeFailure)
      throws Exception {
    var request =
        request(
            List.of(
                ensureData(),
                new MutationStep(
                    new StepId("missing-sheet"),
                    new SheetSelector.ByName(new SheetName("Missing")),
                    new CellMutationAction.AppendRow(
                        new CellRowSpec.Typed(
                            List.of(new CellValueSpec.Text(TextSourceSpec.inline("one"))))))));
    var normal = execute(request, CloseFailure.NONE);
    var failedClose = execute(request, closeFailure);
    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, normal.problem().code());
    assertPrimaryAndFactsRetained(normal, failedClose, new StepId("missing-sheet"), closeFailure);
  }

  @ParameterizedTest
  @EnumSource(
      value = CloseFailure.class,
      names = {"IO", "UNCHECKED"})
  void retainsCollectedAssertionFailureAndCompletedObservationWhenCloseAlsoFails(
      CloseFailure closeFailure) throws Exception {
    var target = new CellSelector.ByAddress(new SheetName("Data"), new CellAddress("A1"));
    var request =
        request(
            List.of(
                ensureData(),
                new InspectionStep(
                    new StepId("observed"), target, new SheetIntrospectionQuery.GetCells()),
                new AssertionStep(
                    new StepId("expected-text"),
                    target,
                    new CellAssertion.Value(new CellValue.Text("expected")))));
    var normal = execute(request, CloseFailure.NONE);
    var failedClose = execute(request, closeFailure);
    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, normal.problem().code());
    assertEquals(1, normal.assertions().size());
    assertInstanceOf(AssertionResult.Failed.class, normal.assertions().getFirst());
    assertEquals(1, normal.inspections().size());
    assertPrimaryAndFactsRetained(normal, failedClose, new StepId("expected-text"), closeFailure);
  }

  private WorkbookResult.Failure execute(WorkbookPlan request, CloseFailure closeFailure)
      throws IOException {
    ExecutionInputBindings bindings = ExecutionInputBindingsFixtureSupport.bindings(root, request);
    AtomicInteger closes = new AtomicInteger();
    AtomicInteger publications = new AtomicInteger();
    WorkbookExecutionEngine engine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectors = new SemanticSelectorResolver(engine);
    TempFileFactory temp = bindings.tempFileFactory();
    StepExecution steps =
        new StepExecution(
            engine, selectors, new AssertionExecutor(engine, selectors), temp, bindings);
    ExecutionStreamingWorkflow workflow =
        new ExecutionStreamingWorkflow(
            (staged, persistence, source, inputs, acceptance) -> {
              publications.incrementAndGet();
              throw new AssertionError("Failed execution must not publish");
            },
            new CalculationWorkflow(writer -> {}),
            steps,
            temp,
            writer -> {
              closes.incrementAndGet();
              writer.close();
              switch (closeFailure) {
                case NONE -> {}
                case IO -> throw new IOException("synthetic writer close failure");
                case UNCHECKED ->
                    throw new IllegalStateException("synthetic unchecked close failure");
              }
            });
    try (var prepared = RequestPreflight.verify(request, bindings).requirePreparedRequest()) {
      var result =
          WorkbookTempScope.call(
              temp::createTempFile,
              () ->
                  workflow.execute(
                      new ExecutionStreamingWorkflow.Call(
                          ProtocolVersion.CURRENT,
                          prepared,
                          new ExecutionModeSpec.StreamingWrite(),
                          List.of(),
                          ExecutionJournalRecorder.start(
                              request, ExecutionProgressSink.NOOP, root))));
      assertEquals(1, closes.get());
      assertEquals(0, publications.get());
      return assertInstanceOf(WorkbookResult.Failure.class, result);
    } finally {
      bindings.closeRunResources();
    }
  }

  private static void assertPrimaryAndFactsRetained(
      WorkbookResult.Failure normal,
      WorkbookResult.Failure failedClose,
      StepId id,
      CloseFailure closeFailure) {
    assertEquals(normal.problem().code(), failedClose.problem().code());
    assertEquals(normal.problem().category(), failedClose.problem().category());
    assertEquals(normal.problem().recovery(), failedClose.problem().recovery());
    assertEquals(normal.problem().title(), failedClose.problem().title());
    assertEquals(normal.problem().message(), failedClose.problem().message());
    assertEquals(normal.problem().resolution(), failedClose.problem().resolution());
    assertEquals(normal.problem().context(), failedClose.problem().context());
    var context =
        assertInstanceOf(ProblemContext.ExecuteStep.class, failedClose.problem().context());
    assertEquals(id, context.step().stepId());
    GridGrindProblemCode expectedCloseCode =
        closeFailure == CloseFailure.IO
            ? GridGrindProblemCode.IO_ERROR
            : GridGrindProblemCode.INTERNAL_ERROR;
    assertTrue(
        failedClose.problem().causes().stream()
            .anyMatch(cause ->
                cause.code() == expectedCloseCode && "EXECUTE_REQUEST".equals(cause.stage())));
    assertEquals(normal.problem().causes().size() + 1, failedClose.problem().causes().size());
    assertEquals(normal.assertions(), failedClose.assertions());
    assertEquals(normal.inspections(), failedClose.inspections());
    assertEquals(normal.calculation(), failedClose.calculation());
    assertEquals(normal.persistence(), failedClose.persistence());
    var normalJournal =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, normal.journal().outcome());
    var failedCloseJournal =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failedClose.journal().outcome());
    assertEquals(normalJournal.completedStepCount(), failedCloseJournal.completedStepCount());
    assertEquals(normalJournal.failedStep(), failedCloseJournal.failedStep());
  }

  private WorkbookPlan request(List<WorkbookStep> steps) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            root.resolve("never-published.xlsx").toString(),
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        new ExecutionPolicySpec(
            new ExecutionModeSpec.StreamingWrite(),
            ExecutionJournalSpec.defaults(),
            CalculationPolicySpec.defaults(),
            AssertionModeSpec.COLLECT),
        FormulaEnvironmentSpec.empty(),
        steps);
  }

  /** Independent native cleanup outcome injected after the real writer closes. */
  private enum CloseFailure {
    NONE,
    IO,
    UNCHECKED
  }

  private static MutationStep ensureData() {
    return new MutationStep(
        new StepId("ensure"),
        new SheetSelector.ByName(new SheetName("Data")),
        new WorkbookMutationAction.EnsureSheet());
  }
}
