package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies streaming publication failures retain their truthful persistence state. */
class ExecutionStreamingWorkflowTest {
  @TempDir Path root;

  @Test
  void mapsAStreamingPublicationFailureToItsKnownNotPublishedOutcome() throws Exception {
    Path output = root.resolve("output.xlsx");
    WorkbookPlan request = streamingRequest(output);
    WorkbookResultPersistence.PublicationOutcome.NotPublished publication =
        new WorkbookResultPersistence.PublicationOutcome.NotPublished(
            WorkbookResultPersistence.PublicationOutcome.DestinationState.ABSENT);
    ExecutionInputBindings bindings = ExecutionInputBindingsFixtureSupport.bindings(root, request);
    ExecutionStreamingWorkflow workflow =
        new ExecutionStreamingWorkflow(
            (staged, persistence, source, preparedInputs, acceptance) -> {
              throw new WorkbookPublicationException(
                  publication, new IOException("publication failed"));
            },
            new CalculationWorkflow(writer -> {}),
            stepSupport(bindings),
            WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile);
    try {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  new ExecutionStreamingWorkflow.Call(
                      ProtocolVersion.CURRENT,
                      RequestPreflight.verify(request, bindings).requirePreparedRequest(),
                      new ExecutionModeSpec.StreamingWrite(),
                      List.of(),
                      ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root))));

      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
      assertEquals(
          publication,
          assertInstanceOf(
                  WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, failure.persistence())
              .publication());
      assertTrue(Files.notExists(output));
    } finally {
      bindings.closeRunResources();
    }
  }

  @Test
  void mapsAnUnexpectedStreamingPersistenceFailureWithoutInventingAPublicationOutcome()
      throws Exception {
    Path output = root.resolve("unexpected-output.xlsx");
    WorkbookPlan request = streamingRequest(output);
    ExecutionInputBindings bindings = ExecutionInputBindingsFixtureSupport.bindings(root, request);
    ExecutionStreamingWorkflow workflow =
        new ExecutionStreamingWorkflow(
            (staged, persistence, source, preparedInputs, acceptance) -> {
              throw new IOException("storage unavailable");
            },
            new CalculationWorkflow(writer -> {}),
            stepSupport(bindings),
            WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile);

    try {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  new ExecutionStreamingWorkflow.Call(
                      ProtocolVersion.CURRENT,
                      RequestPreflight.verify(request, bindings).requirePreparedRequest(),
                      new ExecutionModeSpec.StreamingWrite(),
                      List.of(),
                      ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root))));

      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
      assertTrue(Files.notExists(output));
    } finally {
      bindings.closeRunResources();
    }
  }

  private WorkbookPlan streamingRequest(Path output) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        ExecutionPolicySpec.mode(new ExecutionModeSpec.StreamingWrite()),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new dev.erst.gridgrind.core.plan.MutationStep(
                new dev.erst.gridgrind.core.vocabulary.StepId("ensure-sheet"),
                new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Data")),
                new dev.erst.gridgrind.core.plan.WorkbookMutationAction.EnsureSheet())));
  }

  private StepExecution stepSupport(ExecutionInputBindings bindings) {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new StepExecution(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        bindings);
  }
}
