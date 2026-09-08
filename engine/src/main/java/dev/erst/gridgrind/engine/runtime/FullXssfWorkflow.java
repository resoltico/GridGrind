package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Orchestrates the one admitted full-XSSF execution path. */
final class FullXssfWorkflow {
  private final ExecutionCalculationSupport calculationSupport;
  private final FullXssfStepLoop stepLoop;
  private final FullXssfPersistence persistence;
  private final FullXssfHostAcceptance hostAcceptance;
  private final ExecutionResponseSupport responseSupport;

  FullXssfWorkflow(
      ExecutionCalculationSupport calculationSupport,
      FullXssfStepExecution stepExecution,
      FullXssfCalculationCheckpoint calculationCheckpoint,
      FullXssfPersistence persistence,
      HostAcceptanceExecutor hostAcceptanceExecutor,
      ExecutionResponseSupport responseSupport) {
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.stepLoop =
        new FullXssfStepLoop(
            Objects.requireNonNull(stepExecution, "stepExecution must not be null"),
            Objects.requireNonNull(
                calculationCheckpoint, "calculationCheckpoint must not be null"));
    this.persistence = Objects.requireNonNull(persistence, "persistence must not be null");
    this.hostAcceptance =
        new FullXssfHostAcceptance(
            Objects.requireNonNull(
                hostAcceptanceExecutor, "hostAcceptanceExecutor must not be null"),
            this.calculationSupport);
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
  }

  FullXssfWorkflow(
      ExecutionCalculationSupport calculationSupport,
      FullXssfStepLoop stepLoop,
      FullXssfPersistence persistence,
      FullXssfHostAcceptance hostAcceptance,
      ExecutionResponseSupport responseSupport) {
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.stepLoop = Objects.requireNonNull(stepLoop, "stepLoop must not be null");
    this.persistence = Objects.requireNonNull(persistence, "persistence must not be null");
    this.hostAcceptance = Objects.requireNonNull(hostAcceptance, "hostAcceptance must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
  }

  WorkbookResult execute(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    FullXssfWorkflowState state =
        new FullXssfWorkflowState(protocolVersion, request, workbook, warnings, journal, bindings);
    FullXssfHostAcceptance.Context acceptance;
    try {
      acceptance = hostAcceptance.capture(state, bindings);
    } catch (IOException exception) {
      return closeHostFailure(state, exception);
    }
    Optional<WorkbookResult> earlyOutcome = stepLoop.execute(state, executionMode);
    if (earlyOutcome.isPresent()) {
      return earlyOutcome.orElseThrow();
    }
    earlyOutcome = completePlanCalculation(state);
    if (earlyOutcome.isPresent()) {
      return earlyOutcome.orElseThrow();
    }
    earlyOutcome = closeCollectedAssertionFailure(state);
    if (earlyOutcome.isPresent()) {
      return earlyOutcome.orElseThrow();
    }
    earlyOutcome = executeRequiredCalculation(state, acceptance);
    if (earlyOutcome.isPresent()) {
      return earlyOutcome.orElseThrow();
    }
    earlyOutcome = verifyHostAcceptance(state, acceptance);
    if (earlyOutcome.isPresent()) {
      return earlyOutcome.orElseThrow();
    }
    return persistAndClose(state, bindings, acceptance);
  }

  private Optional<WorkbookResult> completePlanCalculation(FullXssfWorkflowState state) {
    if (state.calculationExecuted()) {
      return Optional.empty();
    }
    ExecutionCalculationSupport.CalculationExecutionOutcome outcome =
        calculationSupport.executeCalculationPolicy(
            state.executionContext().workbook(),
            state.executionContext().request(),
            state.executionContext().journal(),
            state.formulaOrigins(),
            null);
    state.calculation(outcome.report());
    state.executionContext().warnings().addAll(outcome.warnings());
    return outcome.failure().map(problem -> closeFailure(state, problem));
  }

  private Optional<WorkbookResult> closeCollectedAssertionFailure(FullXssfWorkflowState state) {
    if (state.collectedAssertionFailures().isEmpty()) {
      return Optional.empty();
    }
    CollectedAssertionFailures.Failure firstFailure = state.collectedAssertionFailures().first();
    return Optional.of(
        responseSupport.closeWorkbook(
            state.executionContext().workbook(),
            ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                state
                    .executionContext()
                    .failure(
                        state.calculation(),
                        firstFailure.problem(),
                        firstFailure.stepIndex(),
                        firstFailure.stepId())),
            state.executionContext().request(),
            state.executionContext().journal(),
            firstFailure.problem().code(),
            firstFailure.stepIndex(),
            firstFailure.stepId()));
  }

  private Optional<WorkbookResult> executeRequiredCalculation(
      FullXssfWorkflowState state, FullXssfHostAcceptance.Context acceptance) {
    return hostAcceptance
        .requireCalculation(state, acceptance)
        .map(problem -> closeFailure(state, problem));
  }

  private Optional<WorkbookResult> verifyHostAcceptance(
      FullXssfWorkflowState state, FullXssfHostAcceptance.Context acceptance) {
    try {
      hostAcceptance.verify(state, acceptance);
      return Optional.empty();
    } catch (AssertionFailedException | IOException exception) {
      hostAcceptance.recordFailure(state, exception);
      return Optional.of(closeHostFailure(state, exception));
    }
  }

  private WorkbookResult persistAndClose(
      FullXssfWorkflowState state,
      ExecutionInputBindings bindings,
      FullXssfHostAcceptance.Context acceptance) {
    FullXssfPersistence.Result outcome =
        persistence.persist(
            state.executionContext(),
            bindings,
            state.calculation(),
            hostAcceptance.stagedArtifactAcceptance(state, acceptance));
    if (outcome.failureResponse() != null) {
      return outcome.failureResponse();
    }
    WorkbookResultPersistence.PersistenceOutcome persistenceOutcome =
        Objects.requireNonNull(outcome.persistence(), "persistence must exist on success");
    return responseSupport.closeWorkbook(
        state.executionContext().workbook(),
        new WorkbookResult.Success(
            state.executionContext().protocolVersion(),
            state.executionContext().request().planId(),
            state
                .executionContext()
                .journal()
                .buildSuccess(state.executionContext().request().steps().size(), false),
            state.calculation(),
            persistenceOutcome,
            ExecutionEvidence.forExecution(
                state.executionContext().request(),
                bindings,
                state.calculation(),
                state.executionContext().planAssertions(),
                state.hostAssertions(),
                state.hostPreservation(),
                !(persistenceOutcome
                    instanceof WorkbookResultPersistence.PersistenceOutcome.NotSaved)),
            state.executionContext().warnings(),
            List.copyOf(state.executionContext().planAssertions()),
            List.copyOf(state.executionContext().inspections())),
        state.executionContext().request(),
        state.executionContext().journal(),
        null,
        null,
        null);
  }

  private WorkbookResult closeFailure(
      FullXssfWorkflowState state, GridGrindProblemDetail.Problem problem) {
    return responseSupport.closeWorkbook(
        state.executionContext().workbook(),
        ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
            state.executionContext().failure(state.calculation(), problem)),
        state.executionContext().request(),
        state.executionContext().journal(),
        problem.code(),
        null,
        null);
  }

  private WorkbookResult closeHostFailure(FullXssfWorkflowState state, Exception exception) {
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            exception,
            new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                ExecutionRequestPaths.requestShape(state.executionContext().request())));
    return closeFailure(state, problem);
  }
}
