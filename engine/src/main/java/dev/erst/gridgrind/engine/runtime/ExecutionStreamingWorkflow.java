package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Streaming-write workflow for mutation-first execution against a transient workbook writer. */
final class ExecutionStreamingWorkflow {
  private final StreamingWorkbookPersistence streamingWorkbookPersistence;
  private final ExecutionCalculationSupport calculationSupport;
  private final ExecutionStepSupport stepSupport;
  private final HostAcceptanceExecutor hostAcceptanceExecutor;
  private final TempFileFactory tempFileFactory;

  ExecutionStreamingWorkflow(
      ExecutionWorkbookSupport workbookSupport,
      ExecutionCalculationSupport calculationSupport,
      ExecutionStepSupport stepSupport,
      TempFileFactory tempFileFactory) {
    this(
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null")
            ::persistStreamingWorkbook,
        calculationSupport,
        stepSupport,
        tempFileFactory);
  }

  ExecutionStreamingWorkflow(
      StreamingWorkbookPersistence streamingWorkbookPersistence,
      ExecutionCalculationSupport calculationSupport,
      ExecutionStepSupport stepSupport,
      TempFileFactory tempFileFactory) {
    this.streamingWorkbookPersistence =
        Objects.requireNonNull(
            streamingWorkbookPersistence, "streamingWorkbookPersistence must not be null");
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.hostAcceptanceExecutor = new HostAcceptanceExecutor(this.stepSupport);
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  WorkbookResult execute(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    CollectedAssertionFailures collectedAssertionFailures = new CollectedAssertionFailures();
    StreamingWorkflowContext workflowContext =
        StreamingWorkflowContext.create(
            protocolVersion, request, executionMode, warnings, journal, bindings);
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    @Nullable Path materializedPath = null;
    boolean movedToPersistenceTarget = false;
    recordOpen(journal);

    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      java.util.Optional<WorkbookResult> stepFailure =
          executeStreamingSteps(writer, workflowContext, calculation, collectedAssertionFailures);
      if (stepFailure.isPresent()) {
        return stepFailure.orElseThrow();
      }

      ExecutionCalculationSupport.CalculationExecutionOutcome calculationOutcome =
          calculationSupport.executeStreamingCalculationPolicy(writer, request, journal);
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure().isPresent()) {
        ExecutionWorkbookSupport.deleteIfExists(materializedPath);
        GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
        return ExecutionResponseSupport.failureResponse(
            workflowContext.failure(calculation, problem));
      }

      if (!collectedAssertionFailures.isEmpty()) {
        CollectedAssertionFailures.Failure firstFailure = collectedAssertionFailures.first();
        return ExecutionResponseSupport.failureResponse(
            workflowContext.failure(
                calculation,
                firstFailure.problem(),
                firstFailure.stepIndex(),
                firstFailure.stepId()));
      }

      materializedPath =
          ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
              tempFileFactory.createTempFile("gridgrind-streaming-write-", ".xlsx"));
      writer.save(materializedPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      HostAcceptanceVerification hostAcceptance =
          hostAcceptanceExecutor.verifyMaterializedTerminalAcceptance(
              ((dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded)
                      bindings.executionGrant())
                  .hostAcceptancePolicy(),
              materializedPath,
              workflowContext.workbookLocation());
      workflowContext.hostAssertions().addAll(hostAcceptance.assertions());
      workflowContext.preservation().addAll(hostAcceptance.preservation());
    } catch (AssertionFailedException exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      workflowContext
          .hostAssertions()
          .add(
              new AssertionResult.Failed(
                  exception.assertionFailure().stepId(),
                  exception.assertionFailure().assertionType(),
                  exception.assertionFailure()));
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return ExecutionResponseSupport.failureResponse(
          workflowContext.failure(calculation, problem));
    } catch (IOException exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return ExecutionResponseSupport.failureResponse(
          workflowContext.failure(calculation, problem));
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    WorkbookResultPersistence.PersistenceOutcome persistence;
    try {
      persistence =
          streamingWorkbookPersistence.persist(
              materializedPath,
              request.persistence(),
              request.source(),
              bindings,
              StagedArtifactAcceptance.none());
      movedToPersistenceTarget =
          !(persistence instanceof WorkbookResultPersistence.PersistenceOutcome.NotSaved);
    } catch (WorkbookPublicationException exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory())));
      persistencePhase.fail(problem.code());
      WorkbookResultPersistence.PersistenceOutcome failedPersistence =
          dev.erst.gridgrind.contract.dto.WorkbookResults.publicationPersistenceOutcome(
              request, exception.publication());
      return ExecutionResponseSupport.failureResponse(
          workflowContext.failure(calculation, problem), failedPersistence);
    } catch (Exception exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory())));
      persistencePhase.fail(problem.code());
      return ExecutionResponseSupport.failureResponse(
          workflowContext.failure(calculation, problem));
    } finally {
      if (!movedToPersistenceTarget) {
        ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      }
    }
    persistencePhase.succeed();

    return StreamingWorkflowResult.success(
        protocolVersion, request, bindings, workflowContext, calculation, persistence);
  }

  private static void recordOpen(ExecutionJournalRecorder journal) {
    journal.beginOpen().succeed();
  }

  /** Persists one verified streaming artifact through the normal execution authority boundary. */
  @FunctionalInterface
  interface StreamingWorkbookPersistence {
    /** Persists one staged streaming artifact and returns the resulting persistence fact. */
    WorkbookResultPersistence.PersistenceOutcome persist(
        Path stagedArtifact,
        WorkbookPlan.WorkbookPersistence persistence,
        WorkbookPlan.WorkbookSource source,
        ExecutionInputBindings bindings,
        StagedArtifactAcceptance stagedArtifactAcceptance)
        throws IOException;
  }

  private java.util.Optional<WorkbookResult> executeStreamingSteps(
      ExcelStreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      CalculationReport calculation,
      CollectedAssertionFailures collectedAssertionFailures) {
    List<WorkbookStep> steps = workflowContext.request().steps();
    for (int stepIndex = 0; stepIndex < steps.size(); stepIndex++) {
      java.util.Optional<WorkbookResult> failure =
          executeStreamingStep(
              writer,
              workflowContext,
              calculation,
              collectedAssertionFailures,
              stepIndex,
              steps.get(stepIndex));
      if (failure.isPresent()) {
        return failure;
      }
    }
    return java.util.Optional.empty();
  }

  private java.util.Optional<WorkbookResult> executeStreamingStep(
      ExcelStreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      CalculationReport calculation,
      CollectedAssertionFailures collectedAssertionFailures,
      int stepIndex,
      WorkbookStep step) {
    ExecutionJournalRecorder.StepHandle stepHandle =
        workflowContext.journal().beginStep(stepIndex, step);
    try {
      java.util.Optional<AssertionFailedException> collectedFailure =
          executeStreamingStep(writer, workflowContext, step);
      recordStreamingStepResult(
          workflowContext,
          collectedAssertionFailures,
          stepIndex,
          step,
          stepHandle,
          collectedFailure);
      return java.util.Optional.empty();
    } catch (Exception exception) {
      return java.util.Optional.of(
          closeFailedStreamingStep(
              workflowContext, calculation, null, stepIndex, step, stepHandle, exception));
    }
  }

  private static void recordStreamingStepResult(
      StreamingWorkflowContext workflowContext,
      CollectedAssertionFailures collectedAssertionFailures,
      int stepIndex,
      WorkbookStep step,
      ExecutionJournalRecorder.StepHandle stepHandle,
      java.util.Optional<AssertionFailedException> collectedFailure) {
    if (collectedFailure.isEmpty()) {
      stepHandle.succeed();
      return;
    }
    AssertionFailedException assertionFailure = collectedFailure.orElseThrow();
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            assertionFailure,
            ExecutionStepContextFactory.contextFor(
                workflowContext.request(), stepIndex, step, assertionFailure));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().stage(), problem.message());
    collectedAssertionFailures.add(stepIndex, step.stepId(), problem);
  }

  private java.util.Optional<AssertionFailedException> executeStreamingStep(
      ExcelStreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      WorkbookStep step)
      throws IOException, AssertionFailedException {
    return switch (step) {
      case MutationStep mutationStep -> {
        stepSupport.mutationStepExecutor().executeStreaming(writer, mutationStep);
        yield java.util.Optional.empty();
      }
      case AssertionStep assertionStep -> {
        if (workflowContext.request().assertionMode() == AssertionModeInput.FAIL_FAST) {
          workflowContext
              .planAssertions()
              .add(
                  stepSupport.executeStreamingAssertionStep(
                      writer, assertionStep, workflowContext.workbookLocation()));
          yield java.util.Optional.empty();
        }
        AssertionStepExecution assertionExecution =
            stepSupport.executeStreamingAssertionStepCollecting(
                writer, assertionStep, workflowContext.workbookLocation());
        workflowContext.planAssertions().add(assertionExecution.result());
        yield switch (assertionExecution) {
          case AssertionStepExecution.Passed _ -> java.util.Optional.empty();
          case AssertionStepExecution.Failed failed -> java.util.Optional.of(failed.failure());
        };
      }
      case InspectionStep inspectionStep -> {
        workflowContext
            .inspections()
            .add(
                stepSupport.executeStreamingInspectionStep(
                    writer,
                    inspectionStep,
                    workflowContext.workbookLocation(),
                    workflowContext.executionMode()));
        yield java.util.Optional.empty();
      }
    };
  }

  private WorkbookResult closeFailedStreamingStep(
      StreamingWorkflowContext workflowContext,
      CalculationReport calculation,
      @Nullable Path materializedPath,
      int stepIndex,
      WorkbookStep step,
      ExecutionJournalRecorder.StepHandle stepHandle,
      Exception exception) {
    ExecutionWorkbookSupport.deleteIfExists(materializedPath);
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            exception,
            ExecutionStepContextFactory.contextFor(
                workflowContext.request(), stepIndex, step, exception));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().stage(), problem.message());
    if (exception instanceof AssertionFailedException assertionFailed) {
      workflowContext
          .planAssertions()
          .add(
              new AssertionResult.Failed(
                  assertionFailed.assertionFailure().stepId(),
                  assertionFailed.assertionFailure().assertionType(),
                  assertionFailed.assertionFailure()));
    }
    return ExecutionResponseSupport.failureResponse(
        workflowContext.failure(calculation, problem, stepIndex, step.stepId()));
  }
}
