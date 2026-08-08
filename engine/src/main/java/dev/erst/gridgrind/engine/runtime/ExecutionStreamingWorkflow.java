package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionOutcome;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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
      Path workingDirectory) {
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), workingDirectory);
    List<AssertionResult> assertions = new ArrayList<>();
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
          executeStreamingStep(writer, workflowContext, step);
          stepHandle.succeed();
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
            protocolVersion,
            journal,
            request,
            calculation,
            warnings,
            List.copyOf(assertions),
            List.copyOf(inspections),
            problem,
            null,
            null);
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
          protocolVersion,
          journal,
          request,
          calculation,
          warnings,
          List.copyOf(assertions),
          List.copyOf(inspections),
          problem,
          null,
          null);
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    WorkbookResultPersistence.PersistenceOutcome persistence;
    try {
      persistence =
          workbookSupport.persistStreamingWorkbook(
              materializedPath, request.persistence(), request.source(), workingDirectory);
      movedToPersistenceTarget =
          !(persistence instanceof WorkbookResultPersistence.PersistenceOutcome.NotSaved);
    } catch (Exception exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(request, workingDirectory)));
      persistencePhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          request,
          calculation,
          warnings,
          List.copyOf(assertions),
          List.copyOf(inspections),
          problem,
          null,
          null);
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

  private void executeStreamingStep(
      ExcelStreamingWorkbookWriter writer,
      StreamingWorkflowContext workflowContext,
      WorkbookStep step)
      throws IOException, AssertionFailedException {
    switch (step) {
      case MutationStep mutationStep ->
          stepSupport.executeStreamingMutationStep(writer, mutationStep);
      case AssertionStep assertionStep ->
          workflowContext
              .assertions()
              .add(
                  stepSupport.executeStreamingAssertionStep(
                      writer, assertionStep, workflowContext.workbookLocation()));
      case InspectionStep inspectionStep ->
          workflowContext
              .inspections()
              .add(
                  stepSupport.executeStreamingInspectionStep(
                      writer,
                      inspectionStep,
                      workflowContext.workbookLocation(),
                      workflowContext.executionMode()));
    }
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
        workflowContext.protocolVersion(),
        workflowContext.journal(),
        workflowContext.request(),
        calculation,
        workflowContext.warnings(),
        List.copyOf(workflowContext.assertions()),
        List.copyOf(workflowContext.inspections()),
        problem,
        stepIndex,
        step.stepId());
  }

  private record StreamingWorkflowContext(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      WorkbookLocation workbookLocation,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {}
}
