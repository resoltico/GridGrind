package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Persists a verified full-XSSF workbook and preserves publication-state truth on failure. */
final class FullXssfPersistence {
  private final FullXssfWorkbookPersistence workbookPersistence;
  private final ExecutionResponseSupport responseSupport;

  FullXssfPersistence(
      ExecutionWorkbookSupport workbookSupport, ExecutionResponseSupport responseSupport) {
    this(
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null")
            ::persistWorkbook,
        responseSupport);
  }

  FullXssfPersistence(
      FullXssfWorkbookPersistence workbookPersistence, ExecutionResponseSupport responseSupport) {
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
              new dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation.Failed(
                  exception.kind(), exception.identifiers()));
      return failure(executionContext, bindings, calculation, persistencePhase, exception, null);
    } catch (WorkbookPublicationException exception) {
      WorkbookResultPersistence.PersistenceOutcome persistence =
          dev.erst.gridgrind.contract.dto.WorkbookResults.publicationPersistenceOutcome(
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
        ExecutionResponseSupport.problemFor(
            exception,
            new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                ExecutionRequestPaths.requestShape(executionContext.request()),
                ExecutionRequestPaths.persistenceReference(
                    executionContext.request(), bindings.workingDirectory())));
    persistencePhase.fail(problem.code());
    WorkbookResult failure =
        persistence == null
            ? ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                executionContext.failure(calculation, problem))
            : ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
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
        dev.erst.gridgrind.excel.ExcelWorkbook workbook,
        WorkbookPlan.WorkbookSource source,
        WorkbookPlan.WorkbookPersistence persistence,
        ExecutionInputBindings bindings,
        StagedArtifactAcceptance stagedArtifactAcceptance)
        throws PreservationFailedException, WorkbookPublicationException, java.io.IOException;
  }
}
