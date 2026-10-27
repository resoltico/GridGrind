package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Persists a verified full-XSSF workbook and preserves publication-state truth on failure. */
final class FullXssfPersistence {
  private final FullXssfWorkbookPersistence workbookPersistence;
  private final ExecutionResults responseSupport;

  FullXssfPersistence(WorkbookLifecycle workbookSupport, ExecutionResults responseSupport) {
    this(
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null")
            ::persistWorkbook,
        responseSupport);
  }

  FullXssfPersistence(
      FullXssfWorkbookPersistence workbookPersistence, ExecutionResults responseSupport) {
    this.workbookPersistence =
        Objects.requireNonNull(workbookPersistence, "workbookPersistence must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
  }

  Result persist(
      WorkbookWorkflowExecutionContext executionContext,
      ExecutionInputBindings bindings,
      CalculationReport calculation,
      StagedArtifactAcceptance stagedArtifactAcceptance) {
    ExecutionJournalRecorder.PhaseHandle persistencePhase =
        executionContext.journal().beginPersistence();
    try {
      WorkbookResultPersistence.PersistenceOutcome persistence =
          workbookPersistence.persist(
              executionContext.workbook(),
              executionContext.request().source(),
              executionContext.request().persistence(),
              bindings,
              stagedArtifactAcceptance);
      persistencePhase.succeed();
      return new Result(persistence, null);
    } catch (PreservationFailedException exception) {
      executionContext
          .preservation()
          .add(
              new dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation.Failed(
                  exception.kind(), exception.identifiers()));
      return failure(executionContext, bindings, calculation, persistencePhase, exception, null);
    } catch (WorkbookPublicationException exception) {
      WorkbookResultPersistence.PersistenceOutcome persistence =
          dev.erst.gridgrind.core.result.WorkbookResults.publicationPersistenceOutcome(
              executionContext.request(), exception.publication());
      return failure(
          executionContext, bindings, calculation, persistencePhase, exception, persistence);
    } catch (Exception exception) {
      return failure(executionContext, bindings, calculation, persistencePhase, exception, null);
    }
  }

  private Result failure(
      WorkbookWorkflowExecutionContext executionContext,
      ExecutionInputBindings bindings,
      CalculationReport calculation,
      ExecutionJournalRecorder.PhaseHandle persistencePhase,
      Exception exception,
      WorkbookResultPersistence.@Nullable PersistenceOutcome persistence) {
    GridGrindProblemDetail.Problem problem =
        GridGrindProblems.fromException(
            exception,
            new dev.erst.gridgrind.core.problem.ProblemContext.PersistWorkbook(
                ExecutionRequestPaths.requestShape(executionContext.request()),
                ExecutionRequestPaths.persistenceReference(
                    executionContext.request(), bindings.workingDirectory())));
    persistencePhase.fail(problem.code());
    WorkbookResult failure =
        persistence == null
            ? ExecutionResults.failureResponseWithoutPlanOutcomeEvent(
                executionContext.failure(calculation, problem))
            : ExecutionResults.failureResponseWithoutPlanOutcomeEvent(
                executionContext.failure(calculation, problem), persistence);
    return new Result(
        null,
        responseSupport.closeWorkbook(
            executionContext.workbook(),
            failure,
            executionContext.request(),
            executionContext.journal(),
            problem.code(),
            null,
            null));
  }

  record Result(
      WorkbookResultPersistence.@Nullable PersistenceOutcome persistence,
      @Nullable WorkbookResult failureResponse) {}

  /** Persists one verified full-XSSF workbook through the normal execution authority boundary. */
  @FunctionalInterface
  interface FullXssfWorkbookPersistence {
    /** Persists one workbook and returns the resulting persistence fact. */
    WorkbookResultPersistence.PersistenceOutcome persist(
        dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook workbook,
        WorkbookPlan.WorkbookSource source,
        WorkbookPlan.WorkbookPersistence persistence,
        ExecutionInputBindings bindings,
        StagedArtifactAcceptance stagedArtifactAcceptance)
        throws PreservationFailedException, WorkbookPublicationException, java.io.IOException;
  }
}
