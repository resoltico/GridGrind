package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionOutcome;
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
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Streaming-write workflow for mutation-first execution against a transient workbook writer. */
final class ExecutionStreamingWorkflow {
  private final ExecutionWorkbookSupport workbookSupport;
  private final ExecutionCalculationSupport calculationSupport;
  private final ExecutionStepSupport stepSupport;
  private final TempFileFactory tempFileFactory;

  ExecutionStreamingWorkflow(
      ExecutionWorkbookSupport workbookSupport,
      ExecutionCalculationSupport calculationSupport,
      ExecutionStepSupport stepSupport,
      TempFileFactory tempFileFactory) {
    this.workbookSupport =
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null");
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
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
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), bindings.workingDirectory());
    List<AssertionResult> assertions = new ArrayList<>();
    CollectedAssertionFailures collectedAssertionFailures = new CollectedAssertionFailures();
    List<InspectionResult> inspections = new ArrayList<>();
    StreamingWorkflowContext workflowContext =
        new StreamingWorkflowContext(
            protocolVersion,
            request,
            executionMode,
            warnings,
            journal,
            workbookLocation,
            assertions,
            inspections);
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    @Nullable Path materializedPath = null;
    boolean movedToPersistenceTarget = false;
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    openPhase.succeed();

    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
        WorkbookStep step = request.steps().get(stepIndex);
        ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
        try {
          java.util.Optional<AssertionFailedException> collectedFailure =
              executeStreamingStep(writer, workflowContext, step);
          if (collectedFailure.isPresent()) {
            AssertionFailedException assertionFailure = collectedFailure.orElseThrow();
            GridGrindProblemDetail.Problem problem =
                ExecutionResponseSupport.problemFor(
                    assertionFailure,
                    stepSupport.executeStepContext(request, stepIndex, step, assertionFailure));
            stepHandle.fail(
                problem.code(), problem.category(), problem.context().stage(), problem.message());
            collectedAssertionFailures.add(stepIndex, step.stepId(), problem);
          } else {
            stepHandle.succeed();
          }
        } catch (Exception exception) {
          return closeFailedStreamingStep(
              workflowContext,
              calculation,
              materializedPath,
              stepIndex,
              step,
              stepHandle,
              exception);
        }
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
          workbookSupport.persistStreamingWorkbook(
              materializedPath, request.persistence(), request.source(), bindings);
      movedToPersistenceTarget =
          !(persistence instanceof WorkbookResultPersistence.PersistenceOutcome.NotSaved);
    } catch (Exception exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory())));
      persistencePhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          workflowContext.failure(calculation, problem));
    } finally {
      if (!movedToPersistenceTarget) {
        ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      }
    }
    persistencePhase.succeed();

    return new WorkbookResult.Success(
        protocolVersion,
        request.planId(),
        journal.buildSuccess(request.steps().size()),
        calculation,
        persistence,
        warnings,
        List.copyOf(assertions),
        List.copyOf(inspections));
  }

  private java.util.Optional<AssertionFailedException> executeStreamingStep(
      ExcelStreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      WorkbookStep step)
      throws IOException, AssertionFailedException {
    return switch (step) {
      case MutationStep mutationStep -> {
        stepSupport.executeStreamingMutationStep(writer, mutationStep);
        yield java.util.Optional.empty();
      }
      case AssertionStep assertionStep -> {
        if (workflowContext.request().assertionMode() == AssertionModeInput.FAIL_FAST) {
          workflowContext
              .assertions()
              .add(
                  stepSupport.executeStreamingAssertionStep(
                      writer, assertionStep, workflowContext.workbookLocation()));
          yield java.util.Optional.empty();
        }
        AssertionStepExecution assertionExecution =
            stepSupport.executeStreamingAssertionStepCollecting(
                writer, assertionStep, workflowContext.workbookLocation());
        workflowContext.assertions().add(assertionExecution.result());
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
            stepSupport.executeStepContext(workflowContext.request(), stepIndex, step, exception));
    stepHandle.fail(
        problem.code(), problem.category(), problem.context().stage(), problem.message());
    if (exception instanceof AssertionFailedException assertionFailed) {
      workflowContext
          .assertions()
          .add(
              new AssertionResult(
                  AssertionOutcome.FAILED,
                  assertionFailed.assertionFailure().stepId(),
                  assertionFailed.assertionFailure().assertionType()));
    }
    return ExecutionResponseSupport.failureResponse(
        workflowContext.failure(calculation, problem, stepIndex, step.stepId()));
  }

  private record StreamingWorkflowContext(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      WorkbookLocation workbookLocation,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    private ExecutionFailure failure(
        CalculationReport calculation, GridGrindProblemDetail.Problem problem) {
      return failure(calculation, problem, null, null);
    }

    private ExecutionFailure failure(
        CalculationReport calculation,
        GridGrindProblemDetail.Problem problem,
        int failedStepIndex,
        String failedStepId) {
      return failure(calculation, problem, Integer.valueOf(failedStepIndex), failedStepId);
    }

    private ExecutionFailure failure(
        CalculationReport calculation,
        GridGrindProblemDetail.Problem problem,
        @Nullable Integer failedStepIndex,
        @Nullable String failedStepId) {
      return new ExecutionFailure(
          new ExecutionFailure.Context(protocolVersion, journal, request, calculation),
          new ExecutionFailure.Artifacts(warnings, assertions, inspections),
          new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId));
    }
  }
}
