package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
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
            ExecutionPolicyInput.mode(new ExecutionModeInput.StreamingWrite()),
            FormulaEnvironmentInput.empty(),
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
            new ExecutionCalculationSupport(writer -> {}),
            stepSupport(),
            WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile);
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            new GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new GridGrindExecutionGrant.PublicationAuthority.None(),
                List.of(),
                GridGrindHostAcceptancePolicy.minimum()));

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            workflow.execute(
                GridGrindProtocolVersion.current(),
                request,
                new ExecutionModeInput.StreamingWrite(),
                List.of(),
                ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                bindings));

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
            new ExecutionCalculationSupport(writer -> {}),
            stepSupport(),
            WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile);

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            workflow.execute(
                GridGrindProtocolVersion.current(),
                request,
                new ExecutionModeInput.StreamingWrite(),
                List.of(),
                ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                bindings()));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertTrue(Files.notExists(output));
  }

  private WorkbookPlan streamingRequest(Path output) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            output.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        ExecutionPolicyInput.mode(new ExecutionModeInput.StreamingWrite()),
        FormulaEnvironmentInput.empty(),
        List.of());
  }

  private ExecutionInputBindings bindings() {
    return new ExecutionInputBindings(
        root,
        root.resolve("scratch"),
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            GridGrindHostAcceptancePolicy.minimum()));
  }

  private ExecutionStepSupport stepSupport() {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new ExecutionStepSupport(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        new ExecutionInputBindings(
            root, root.resolve("scratch"), ExecutionGrantTestSupport.noPublication()));
  }
}
