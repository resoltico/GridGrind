package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.engine.backend.ExecutionBackend;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactIo;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Direct event-read workflow for inspection-only execution against an existing workbook file. */
final class ExecutionDirectEventReadWorkflow
    implements ExecutionBackend<ExecutionDirectEventReadWorkflow.Call> {
  private final StepExecution stepSupport;
  private final MaterializedTerminalAcceptance terminalAcceptance;
  private final ExecutionResults responseSupport;
  private final TempFileFactory tempFileFactory;

  ExecutionDirectEventReadWorkflow(
      StepExecution stepSupport,
      ExecutionResults responseSupport,
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
      StepExecution stepSupport,
      MaterializedTerminalAcceptance terminalAcceptance,
      ExecutionResults responseSupport,
      TempFileFactory tempFileFactory) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.terminalAcceptance =
        Objects.requireNonNull(terminalAcceptance, "terminalAcceptance must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  record Call(
      ProtocolVersion protocolVersion,
      ResolvedPlan plan,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {}

  @Override
  public BackendCapabilities.Mode capabilities() {
    return BackendCapabilities.eventRead();
  }

  @Override
  public WorkbookResult execute(Call call) {
    ProtocolVersion protocolVersion = call.protocolVersion();
    ResolvedPlan plan = call.plan();
    List<RequestWarning> warnings = call.warnings();
    ExecutionJournalRecorder journal = call.journal();
    ExecutionInputBindings bindings = call.bindings();

    WorkbookPlan request = plan.request();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    WorkbookPlan.WorkbookSource.ExistingFile source =
        (WorkbookPlan.WorkbookSource.ExistingFile) request.source();
    List<InspectionResult> inspections = new ArrayList<>();
    List<AssertionResult> hostAssertions = new ArrayList<>();
    List<dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation> hostPreservation =
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
              ExecutionRequestPaths.materializeOpenSecurity(
                  source.security().orElse(null), bindings),
              tempFileFactory::createTempFile);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.OpenWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory())));
      openPhase.fail(problem.code());
      return responseSupport.closeReadableWorkbook(
          null,
          ExecutionResults.failureResponseWithoutPlanOutcomeEvent(
              executionContext.failure(problem)),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
    openPhase.succeed();
    dev.erst.gridgrind.core.grant.AcceptancePolicy hostAcceptancePolicy =
        bindings.executionGrant().acceptancePolicy();
    for (int stepIndex = 0; stepIndex < plan.steps().size(); stepIndex++) {
      InspectionStep inspectionStep = (InspectionStep) plan.steps().get(stepIndex);
      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, inspectionStep);
      try {
        inspections.add(
            stepSupport.executeEventInspection(materialized.workbookPath(), inspectionStep));
        stepHandle.succeed();
      } catch (Exception exception) {
        GridGrindProblemDetail.Problem problem =
            GridGrindProblems.fromException(
                exception,
                ExecutionStepContextFactory.contextFor(
                    request, stepIndex, inspectionStep, exception));
        stepHandle.fail(
            problem.code(), problem.category(), problem.context().type(), problem.message());
        return responseSupport.closeReadableWorkbook(
            materialized,
            ExecutionResults.failureResponseWithoutPlanOutcomeEvent(
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
        GridGrindProblems.fromException(
            exception,
            new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                ExecutionRequestPaths.requestShape(executionContext.request())));
    return responseSupport.closeReadableWorkbook(
        materialized,
        ExecutionResults.failureResponseWithoutPlanOutcomeEvent(executionContext.failure(problem)),
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
        dev.erst.gridgrind.core.grant.AcceptancePolicy policy,
        java.nio.file.Path workbookPath,
        dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation workbookLocation)
        throws IOException, AssertionFailedException;
  }
}
