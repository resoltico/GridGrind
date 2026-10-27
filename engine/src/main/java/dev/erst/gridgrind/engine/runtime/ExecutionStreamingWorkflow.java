package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionModeSpec;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.PreparedStep;
import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.engine.backend.ExecutionBackend;
import dev.erst.gridgrind.engine.backend.streaming.StreamingWorkbookWriter;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Streaming-write workflow for mutation-first execution against a transient workbook writer. */
final class ExecutionStreamingWorkflow
    implements ExecutionBackend<ExecutionStreamingWorkflow.Call> {
  private final StreamingWorkbookPersistence streamingWorkbookPersistence;
  private final CalculationWorkflow calculationSupport;
  private final StepExecution stepSupport;
  private final HostAcceptanceExecutor hostAcceptanceExecutor;
  private final TempFileFactory tempFileFactory;

  ExecutionStreamingWorkflow(
      WorkbookLifecycle workbookSupport,
      CalculationWorkflow calculationSupport,
      StepExecution stepSupport,
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
      CalculationWorkflow calculationSupport,
      StepExecution stepSupport,
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

  record Call(
      ProtocolVersion protocolVersion,
      ResolvedPlan plan,
      ExecutionModeSpec executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {}

  @Override
  public BackendCapabilities.Mode capabilities() {
    return BackendCapabilities.streamingWrite();
  }

  @Override
  public WorkbookResult execute(Call call) {

    WorkbookPlan request = call.plan().request();
    CollectedAssertionFailures collectedAssertionFailures = new CollectedAssertionFailures();
    StreamingWorkflowContext workflowContext =
        StreamingWorkflowContext.create(
            call.protocolVersion(),
            call.plan(),
            call.executionMode(),
            call.warnings(),
            call.journal(),
            call.bindings());
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    @Nullable Path materializedPath = null;
    boolean movedToPersistenceTarget = false;
    recordOpen(call.journal());

    try (StreamingWorkbookWriter writer = new StreamingWorkbookWriter()) {
      java.util.Optional<WorkbookResult> stepFailure =
          executeStreamingSteps(writer, workflowContext, calculation, collectedAssertionFailures);
      if (stepFailure.isPresent()) {
        return stepFailure.orElseThrow();
      }

      CalculationWorkflow.CalculationExecutionOutcome calculationOutcome =
          calculationSupport.executeStreamingCalculationPolicy(writer, request, call.journal());
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure().isPresent()) {
        WorkbookLifecycle.deleteIfExists(materializedPath);
        GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
        return ExecutionResults.failureResponse(workflowContext.failure(calculation, problem));
      }

      if (!collectedAssertionFailures.isEmpty()) {
        CollectedAssertionFailures.Failure firstFailure = collectedAssertionFailures.first();
        return ExecutionResults.failureResponse(
            workflowContext.failure(
                calculation,
                firstFailure.problem(),
                firstFailure.stepIndex(),
                firstFailure.stepId()));
      }

      materializedPath =
          WorkbookTempFileFactory.prepareCreateNewTarget(
              tempFileFactory.createTempFile("gridgrind-streaming-write-", ".xlsx"));
      writer.save(materializedPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      HostAcceptanceVerification hostAcceptance =
          hostAcceptanceExecutor.verifyMaterializedTerminalAcceptance(
              call.bindings().executionGrant().acceptancePolicy(),
              materializedPath,
              workflowContext.workbookLocation());
      workflowContext.hostAssertions().addAll(hostAcceptance.assertions());
      workflowContext.preservation().addAll(hostAcceptance.preservation());
    } catch (AssertionFailedException exception) {
      WorkbookLifecycle.deleteIfExists(materializedPath);
      workflowContext
          .hostAssertions()
          .add(
              new AssertionResult.Failed(
                  exception.assertionFailure().stepId(),
                  exception.assertionFailure().assertionType(),
                  exception.assertionFailure()));
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return ExecutionResults.failureResponse(workflowContext.failure(calculation, problem));
    } catch (IOException exception) {
      WorkbookLifecycle.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return ExecutionResults.failureResponse(workflowContext.failure(calculation, problem));
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = call.journal().beginPersistence();
    WorkbookResultPersistence.PersistenceOutcome persistence;
    try {
      persistence =
          streamingWorkbookPersistence.persist(
              materializedPath,
              request.persistence(),
              request.source(),
              call.bindings(),
              StagedArtifactAcceptance.none());
      movedToPersistenceTarget =
          !(persistence instanceof WorkbookResultPersistence.PersistenceOutcome.NotSaved);
    } catch (WorkbookPublicationException exception) {
      WorkbookLifecycle.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, call.bindings().workingDirectory())));
      persistencePhase.fail(problem.code());
      WorkbookResultPersistence.PersistenceOutcome failedPersistence =
          dev.erst.gridgrind.core.result.WorkbookResults.publicationPersistenceOutcome(
              request, exception.publication());
      return ExecutionResults.failureResponse(
          workflowContext.failure(calculation, problem), failedPersistence);
    } catch (Exception exception) {
      WorkbookLifecycle.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, call.bindings().workingDirectory())));
      persistencePhase.fail(problem.code());
      return ExecutionResults.failureResponse(workflowContext.failure(calculation, problem));
    } finally {
      if (!movedToPersistenceTarget) {
        WorkbookLifecycle.deleteIfExists(materializedPath);
      }
    }
    persistencePhase.succeed();

    return StreamingWorkflowResult.success(
        call.protocolVersion(),
        request,
        call.bindings(),
        workflowContext,
        calculation,
        persistence);
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
      StreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      CalculationReport calculation,
      CollectedAssertionFailures collectedAssertionFailures) {
    List<PreparedStep> steps = workflowContext.plan().steps();
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
      StreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      CalculationReport calculation,
      CollectedAssertionFailures collectedAssertionFailures,
      int stepIndex,
      PreparedStep prepared) {
    WorkbookStep step = prepared.authoredStep();
    ExecutionJournalRecorder.StepHandle stepHandle =
        workflowContext.journal().beginStep(stepIndex, step);
    try {
      java.util.Optional<AssertionFailedException> collectedFailure =
          executeStreamingStep(writer, workflowContext, prepared);
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
        GridGrindProblems.fromException(
            assertionFailure,
            ExecutionStepContextFactory.contextFor(
                workflowContext.request(), stepIndex, step, assertionFailure));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().type(), problem.message());
    collectedAssertionFailures.add(stepIndex, step.stepId(), problem);
  }

  private java.util.Optional<AssertionFailedException> executeStreamingStep(
      StreamingWorkbookWriter writer, StreamingWorkflowContext workflowContext, PreparedStep step)
      throws IOException, AssertionFailedException {
    return switch (step) {
      case PreparedMutationStep mutationStep -> {
        stepSupport.mutationStepExecutor().executeStreaming(writer, mutationStep);
        yield java.util.Optional.empty();
      }
      case AssertionStep assertionStep -> {
        if (workflowContext.request().assertionMode() == AssertionModeSpec.FAIL_FAST) {
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
    WorkbookLifecycle.deleteIfExists(materializedPath);
    GridGrindProblemDetail.Problem problem =
        GridGrindProblems.fromException(
            exception,
            ExecutionStepContextFactory.contextFor(
                workflowContext.request(), stepIndex, step, exception));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().type(), problem.message());
    if (exception instanceof AssertionFailedException assertionFailed) {
      workflowContext
          .planAssertions()
          .add(
              new AssertionResult.Failed(
                  assertionFailed.assertionFailure().stepId(),
                  assertionFailed.assertionFailure().assertionType(),
                  assertionFailed.assertionFailure()));
    }
    return ExecutionResults.failureResponse(
        workflowContext.failure(calculation, problem, stepIndex, step.stepId()));
  }
}
