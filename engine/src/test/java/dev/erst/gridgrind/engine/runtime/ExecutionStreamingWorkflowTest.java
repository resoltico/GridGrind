package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
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
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                output.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicySpec.mode(new ExecutionModeSpec.StreamingWrite()),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookResultPersistence.PublicationOutcome.NotPublished publication =
        new WorkbookResultPersistence.PublicationOutcome.NotPublished(
            WorkbookResultPersistence.PublicationOutcome.DestinationState.ABSENT);
    ExecutionStreamingWorkflow workflow =
        new ExecutionStreamingWorkflow(
            (staged, persistence, source, bindings, acceptance) -> {
              throw new WorkbookPublicationException(
                  publication, new IOException("publication failed"));
            },
            new CalculationWorkflow(writer -> {}),
            stepSupport(),
            WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile);
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            new ExecutionGrant(
                List.of(),
                List.of(),
                new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                AcceptancePolicy.minimum()));

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            workflow.execute(
                new ExecutionStreamingWorkflow.Call(
                    ProtocolVersion.CURRENT,
                    MutationResolution.preparePlan(request),
                    new ExecutionModeSpec.StreamingWrite(),
                    List.of(),
                    ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                    bindings)));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals(
        publication,
        assertInstanceOf(
                WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, failure.persistence())
            .publication());
    assertTrue(Files.notExists(output));
  }

  @Test
  void mapsAnUnexpectedStreamingPersistenceFailureWithoutInventingAPublicationOutcome()
      throws Exception {
    Path output = root.resolve("unexpected-output.xlsx");
    WorkbookPlan request = streamingRequest(output);
    ExecutionStreamingWorkflow workflow =
        new ExecutionStreamingWorkflow(
            (staged, persistence, source, bindings, acceptance) -> {
              throw new IOException("storage unavailable");
            },
            new CalculationWorkflow(writer -> {}),
            stepSupport(),
            WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile);

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            workflow.execute(
                new ExecutionStreamingWorkflow.Call(
                    ProtocolVersion.CURRENT,
                    MutationResolution.preparePlan(request),
                    new ExecutionModeSpec.StreamingWrite(),
                    List.of(),
                    ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                    bindings())));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertTrue(Files.notExists(output));
  }

  private WorkbookPlan streamingRequest(Path output) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        ExecutionPolicySpec.mode(new ExecutionModeSpec.StreamingWrite()),
        FormulaEnvironmentSpec.empty(),
        List.of());
  }

  private ExecutionInputBindings bindings() {
    return new ExecutionInputBindings(
        root,
        root.resolve("scratch"),
        new ExecutionGrant(
            List.of(),
            List.of(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum()));
  }

  private StepExecution stepSupport() {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new StepExecution(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication()));
  }
}
