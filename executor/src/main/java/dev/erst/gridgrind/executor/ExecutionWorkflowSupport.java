package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Request workflow implementations for workbook, direct-event, and streaming execution modes. */
final class ExecutionWorkflowSupport {
  private final ExecutionWorkbookSupport workbookSupport;
  private final ExecutionCalculationSupport calculationSupport;
  private final ExecutionStepSupport stepSupport;
  private final ExecutionResponseSupport responseSupport;
  private final TempFileFactory tempFileFactory;

  ExecutionWorkflowSupport(
      ExecutionWorkbookSupport workbookSupport,
      ExecutionCalculationSupport calculationSupport,
      ExecutionStepSupport stepSupport,
      ExecutionResponseSupport responseSupport,
      TempFileFactory tempFileFactory) {
    this.workbookSupport =
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null");
    this.calculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  GridGrindResponse executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionModeSelection executionModes,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      Path workingDirectory) {
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), workingDirectory);
    List<AssertionResult> assertions = new ArrayList<>();
    List<InspectionResult> inspections = new ArrayList<>();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    boolean calculationExecuted = false;

    for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
      WorkbookStep step = request.steps().get(stepIndex);
      if (!calculationExecuted && shouldExecuteCalculationBeforeStep(request, step)) {
        ExecutionCalculationSupport.CalculationExecutionOutcome calculationOutcome =
            calculationSupport.executeCalculationPolicy(workbook, request, journal);
        calculation = calculationOutcome.report();
        calculationExecuted = true;
        if (calculationOutcome.failure().isPresent()) {
          GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
          return responseSupport.closeWorkbook(
              workbook,
              ExecutionResponseSupport.failureResponse(
                  protocolVersion,
                  journal,
                  request.steps().size(),
                  calculation,
                  problem,
                  null,
                  null),
              request,
              journal,
              problem.code(),
              null,
              null);
        }
      }

      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
      try {
        switch (step) {
          case MutationStep mutationStep -> stepSupport.executeMutationStep(workbook, mutationStep);
          case AssertionStep assertionStep ->
              assertions.add(
                  stepSupport.executeAssertionStep(
                      assertionStep, workbook, workbookLocation, executionModes.readMode()));
          case InspectionStep inspectionStep ->
              inspections.add(
                  stepSupport.executeInspectionStep(
                      inspectionStep, workbook, workbookLocation, executionModes.readMode()));
        }
        stepHandle.succeed();
      } catch (Exception exception) {
        GridGrindProblemDetail.Problem problem =
            ExecutionResponseSupport.problemFor(
                exception, stepSupport.executeStepContext(request, stepIndex, step, exception));
        stepHandle.fail(
            problem.code(), problem.category(), problem.context().stage(), problem.message());
        return responseSupport.closeWorkbook(
            workbook,
            ExecutionResponseSupport.failureResponse(
                protocolVersion,
                journal,
                request.steps().size(),
                calculation,
                problem,
                stepIndex,
                step.stepId()),
            request,
            journal,
            problem.code(),
            stepIndex,
            step.stepId());
      }
    }

    if (!calculationExecuted) {
      ExecutionCalculationSupport.CalculationExecutionOutcome calculationOutcome =
          calculationSupport.executeCalculationPolicy(workbook, request, journal);
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure().isPresent()) {
        GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
        return responseSupport.closeWorkbook(
            workbook,
            ExecutionResponseSupport.failureResponse(
                protocolVersion, journal, request.steps().size(), calculation, problem, null, null),
            request,
            journal,
            problem.code(),
            null,
            null);
      }
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    GridGrindResponsePersistence.PersistenceOutcome persistence;
    try {
      persistence =
          workbookSupport.persistWorkbook(
              workbook, request.source(), request.persistence(), workingDirectory);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(request, workingDirectory)));
      persistencePhase.fail("failed (" + problem.code() + ")");
      return responseSupport.closeWorkbook(
          workbook,
          ExecutionResponseSupport.failureResponse(
              protocolVersion, journal, request.steps().size(), calculation, problem, null, null),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
    persistencePhase.succeed();

    return responseSupport.closeWorkbook(
        workbook,
        new GridGrindResponse.Success(
            protocolVersion,
            journal.buildSuccess(request.steps().size()),
            calculation,
            persistence,
            warnings,
            List.copyOf(assertions),
            List.copyOf(inspections)),
        request,
        journal,
        null,
        null,
        null);
  }

  GridGrindResponse executeDirectEventReadWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      Path workingDirectory) {
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    WorkbookPlan.WorkbookSource.ExistingFile source =
        (WorkbookPlan.WorkbookSource.ExistingFile) request.source();
    List<InspectionResult> inspections = new ArrayList<>();
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    ExcelOoxmlPackageSecuritySupport.ReadableWorkbook materialized;
    try {
      materialized =
          ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
              ExecutionRequestPaths.normalizePath(source.path(), workingDirectory),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(source.security().orElse(null)),
              tempFileFactory::createTempFile);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.workbookReference(request, workingDirectory)));
      openPhase.fail("failed (" + problem.code() + ")");
      return responseSupport.closeReadableWorkbook(
          null,
          ExecutionResponseSupport.failureResponse(
              protocolVersion, journal, request.steps().size(), calculation, problem, null, null),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
    openPhase.succeed();

    for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
      InspectionStep inspectionStep = (InspectionStep) request.steps().get(stepIndex);
      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, inspectionStep);
      try {
        inspections.add(
            stepSupport.executeEventInspection(materialized.workbookPath(), inspectionStep));
        stepHandle.succeed();
      } catch (Exception exception) {
        GridGrindProblemDetail.Problem problem =
            ExecutionResponseSupport.problemFor(
                exception,
                stepSupport.executeStepContext(request, stepIndex, inspectionStep, exception));
        stepHandle.fail(
            problem.code(), problem.category(), problem.context().stage(), problem.message());
        return responseSupport.closeReadableWorkbook(
            materialized,
            ExecutionResponseSupport.failureResponse(
                protocolVersion,
                journal,
                request.steps().size(),
                calculation,
                problem,
                stepIndex,
                inspectionStep.stepId()),
            request,
            journal,
            problem.code(),
            stepIndex,
            inspectionStep.stepId());
      }
    }

    return responseSupport.closeReadableWorkbook(
        materialized,
        new GridGrindResponse.Success(
            protocolVersion,
            journal.buildSuccess(request.steps().size()),
            calculation,
            new GridGrindResponsePersistence.PersistenceOutcome.NotSaved(),
            warnings,
            List.of(),
            List.copyOf(inspections)),
        request,
        journal,
        null,
        null,
        null);
  }

  GridGrindResponse executeStreamingWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeSelection executionModes,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      Path workingDirectory) {
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), workingDirectory);
    List<AssertionResult> assertions = new ArrayList<>();
    List<InspectionResult> inspections = new ArrayList<>();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    Path materializedPath = null;
    boolean movedToPersistenceTarget = false;
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    openPhase.succeed();

    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
        WorkbookStep step = request.steps().get(stepIndex);
        ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
        try {
          switch (step) {
            case MutationStep mutationStep ->
                stepSupport.executeStreamingMutationStep(writer, mutationStep);
            case AssertionStep assertionStep ->
                assertions.add(
                    stepSupport.executeStreamingAssertionStep(
                        writer, assertionStep, workbookLocation));
            case InspectionStep inspectionStep ->
                inspections.add(
                    stepSupport.executeStreamingInspectionStep(
                        writer, inspectionStep, workbookLocation, executionModes.readMode()));
          }
          stepHandle.succeed();
        } catch (Exception exception) {
          ExecutionWorkbookSupport.deleteIfExists(materializedPath);
          GridGrindProblemDetail.Problem problem =
              ExecutionResponseSupport.problemFor(
                  exception, stepSupport.executeStepContext(request, stepIndex, step, exception));
          stepHandle.fail(
              problem.code(), problem.category(), problem.context().stage(), problem.message());
          return ExecutionResponseSupport.failureResponse(
              protocolVersion,
              journal,
              request.steps().size(),
              calculation,
              problem,
              stepIndex,
              step.stepId());
        }
      }

      ExecutionCalculationSupport.CalculationExecutionOutcome calculationOutcome =
          calculationSupport.executeStreamingCalculationPolicy(writer, request, journal);
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure().isPresent()) {
        ExecutionWorkbookSupport.deleteIfExists(materializedPath);
        GridGrindProblemDetail.Problem problem = calculationOutcome.failure().orElseThrow();
        return ExecutionResponseSupport.failureResponse(
            protocolVersion, journal, request.steps().size(), calculation, problem, null, null);
      }

      materializedPath = tempFileFactory.createTempFile("gridgrind-streaming-write-", ".xlsx");
      writer.save(materializedPath);
    } catch (IOException exception) {
      ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                  ExecutionRequestPaths.requestShape(request)));
      return ExecutionResponseSupport.failureResponse(
          protocolVersion, journal, request.steps().size(), calculation, problem, null, null);
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    GridGrindResponsePersistence.PersistenceOutcome persistence;
    try {
      persistence =
          workbookSupport.persistStreamingWorkbook(
              materializedPath, request.persistence(), request.source(), workingDirectory);
      movedToPersistenceTarget =
          !(persistence instanceof GridGrindResponsePersistence.PersistenceOutcome.NotSaved);
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
          protocolVersion, journal, request.steps().size(), calculation, problem, null, null);
    } finally {
      if (!movedToPersistenceTarget) {
        ExecutionWorkbookSupport.deleteIfExists(materializedPath);
      }
    }
    persistencePhase.succeed();

    return new GridGrindResponse.Success(
        protocolVersion,
        journal.buildSuccess(request.steps().size()),
        calculation,
        persistence,
        warnings,
        List.copyOf(assertions),
        List.copyOf(inspections));
  }

  private static boolean shouldExecuteCalculationBeforeStep(
      WorkbookPlan request, WorkbookStep step) {
    return CalculationPolicyExecutor.requiresMutationPrefix(request.calculationPolicy())
        && !(step instanceof MutationStep);
  }
}
