package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Direct event-read workflow for inspection-only execution against an existing workbook file. */
final class ExecutionDirectEventReadWorkflow {
  private final ExecutionStepSupport stepSupport;
  private final MaterializedTerminalAcceptance terminalAcceptance;
  private final ExecutionResponseSupport responseSupport;
  private final TempFileFactory tempFileFactory;

  ExecutionDirectEventReadWorkflow(
      ExecutionStepSupport stepSupport,
      ExecutionResponseSupport responseSupport,
      TempFileFactory tempFileFactory) {
    this(
        stepSupport,
        new HostAcceptanceExecutor(
                Objects.requireNonNull(stepSupport, "stepSupport must not be null"))
            ::verifyMaterializedTerminalAcceptance,
        responseSupport,
        tempFileFactory);
  }

  ExecutionDirectEventReadWorkflow(
      ExecutionStepSupport stepSupport,
      MaterializedTerminalAcceptance terminalAcceptance,
      ExecutionResponseSupport responseSupport,
      TempFileFactory tempFileFactory) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.terminalAcceptance =
        Objects.requireNonNull(terminalAcceptance, "terminalAcceptance must not be null");
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
      ExecutionInputBindings bindings) {
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    WorkbookPlan.WorkbookSource.ExistingFile source =
        (WorkbookPlan.WorkbookSource.ExistingFile) request.source();
    List<InspectionResult> inspections = new ArrayList<>();
    List<AssertionResult> hostAssertions = new ArrayList<>();
    List<dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation> hostPreservation =
        new ArrayList<>();
    DirectEventReadContext executionContext =
        new DirectEventReadContext(
            protocolVersion,
            request,
            warnings,
            journal,
            calculation,
            hostAssertions,
            hostPreservation,
            inspections);
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    WorkbookArtifactIo.MaterializedWorkbook materialized;
    try {
      materialized =
          WorkbookArtifactIo.materializeWorkbook(
              bindings
                  .requestPathAccess()
                  .materializeRead(source.path(), "source", "gridgrind-source-workbook-", ".xlsx"),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(
                  source.security().orElse(null), bindings),
              tempFileFactory::createTempFile);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory())));
      openPhase.fail(problem.code());
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
    dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy hostAcceptancePolicy =
        ((dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded) bindings.executionGrant())
            .hostAcceptancePolicy();
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
                ExecutionStepContextFactory.contextFor(
                    request, stepIndex, inspectionStep, exception));
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

    try {
      HostAcceptanceVerification hostAcceptance =
          terminalAcceptance.verify(
              hostAcceptancePolicy,
              materialized.workbookPath(),
              ExecutionRequestPaths.workbookLocationFor(
                  request.source(), request.persistence(), bindings.workingDirectory()));
      hostAssertions.addAll(hostAcceptance.assertions());
      hostPreservation.addAll(hostAcceptance.preservation());
    } catch (AssertionFailedException exception) {
      hostAssertions.add(
          new AssertionResult.Failed(
              exception.assertionFailure().stepId(),
              exception.assertionFailure().assertionType(),
              exception.assertionFailure()));
      return hostAcceptanceFailure(materialized, executionContext, exception);
    } catch (IOException exception) {
      return hostAcceptanceFailure(materialized, executionContext, exception);
    }

    return responseSupport.closeReadableWorkbook(
        materialized,
        DirectEventReadResult.success(executionContext, bindings),
        request,
        journal,
        null,
        null,
        null);
  }

  private WorkbookResult hostAcceptanceFailure(
      WorkbookArtifactIo.MaterializedWorkbook materialized,
      DirectEventReadContext executionContext,
      Exception exception) {
    GridGrindProblemDetail.Problem problem =
        ExecutionResponseSupport.problemFor(
            exception,
            new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                ExecutionRequestPaths.requestShape(executionContext.request())));
    return responseSupport.closeReadableWorkbook(
        materialized,
        ExecutionResponseSupport.failureResponseWithoutPlanOutcomeEvent(
            executionContext.failure(problem)),
        executionContext.request(),
        executionContext.journal(),
        problem.code(),
        null,
        null);
  }

  /** Verifies terminal host acceptance against one materialized event-read workbook. */
  @FunctionalInterface
  interface MaterializedTerminalAcceptance {
    /** Returns established host acceptance facts or fails before a success response is emitted. */
    HostAcceptanceVerification verify(
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy policy,
        java.nio.file.Path workbookPath,
        dev.erst.gridgrind.excel.WorkbookLocation workbookLocation)
        throws IOException, AssertionFailedException;
  }
}
