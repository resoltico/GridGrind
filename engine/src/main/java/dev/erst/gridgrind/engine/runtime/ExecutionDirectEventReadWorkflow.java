package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Direct event-read workflow for inspection-only execution against an existing workbook file. */
final class ExecutionDirectEventReadWorkflow {
  private final ExecutionStepSupport stepSupport;
  private final ExecutionResponseSupport responseSupport;
  private final TempFileFactory tempFileFactory;

  ExecutionDirectEventReadWorkflow(
      ExecutionStepSupport stepSupport,
      ExecutionResponseSupport responseSupport,
      TempFileFactory tempFileFactory) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  WorkbookResult execute(
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
    DirectEventReadContext executionContext =
        new DirectEventReadContext(
            protocolVersion, request, warnings, journal, calculation, inspections);
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    WorkbookArtifactIo.MaterializedWorkbook materialized;
    try {
      materialized =
          WorkbookArtifactIo.materializeWorkbook(
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
          ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
              executionContext.failure(problem)),
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
            ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
                executionContext.failure(problem, stepIndex, inspectionStep.stepId())),
            request,
            journal,
            problem.code(),
            stepIndex,
            inspectionStep.stepId());
      }
    }

    return responseSupport.closeReadableWorkbook(
        materialized,
        new WorkbookResult.Success(
            protocolVersion,
            request.planId(),
            journal.buildSuccess(request.steps().size(), false),
            calculation,
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            warnings,
            List.of(),
            List.copyOf(inspections)),
        request,
        journal,
        null,
        null,
        null);
  }

  private record DirectEventReadContext(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      CalculationReport calculation,
      List<InspectionResult> inspections) {
    private ExecutionFailure failure(GridGrindProblemDetail.Problem problem) {
      return failure(problem, null, null);
    }

    private ExecutionFailure failure(
        GridGrindProblemDetail.Problem problem, int failedStepIndex, String failedStepId) {
      return failure(problem, Integer.valueOf(failedStepIndex), failedStepId);
    }

    private ExecutionFailure failure(
        GridGrindProblemDetail.Problem problem,
        @Nullable Integer failedStepIndex,
        @Nullable String failedStepId) {
      return new ExecutionFailure(
          new ExecutionFailure.Context(protocolVersion, journal, request, calculation),
          new ExecutionFailure.Artifacts(warnings, List.of(), inspections),
          new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId));
    }
  }
}
