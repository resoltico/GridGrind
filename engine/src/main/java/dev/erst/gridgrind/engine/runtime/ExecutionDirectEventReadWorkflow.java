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
import java.util.List;
import java.util.Objects;

/** Direct event-read workflow for inspection-only execution against an existing workbook file. */
final class ExecutionDirectEventReadWorkflow
    implements ExecutionBackend<ExecutionDirectEventReadWorkflow.Call> {
  private final StepExecution stepSupport;
  private final MaterializedTerminalAcceptance terminalAcceptance;
  private final ExecutionResults responseSupport;

  ExecutionDirectEventReadWorkflow(StepExecution stepSupport, ExecutionResults responseSupport) {
    this(
        stepSupport,
        new HostAcceptanceExecutor(
                Objects.requireNonNull(stepSupport, "stepSupport must not be null"))
            ::verifyMaterializedTerminalAcceptance,
        responseSupport);
  }

  ExecutionDirectEventReadWorkflow(
      StepExecution stepSupport,
      MaterializedTerminalAcceptance terminalAcceptance,
      ExecutionResults responseSupport) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.terminalAcceptance =
        Objects.requireNonNull(terminalAcceptance, "terminalAcceptance must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
  }

  record Call(
      ProtocolVersion protocolVersion,
      RequestPreflight.PreparedRequest prepared,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal) {
    Call {
      Objects.requireNonNull(prepared, "prepared must not be null");
    }

    ResolvedPlan plan() {
      return prepared.plan();
    }

    ExecutionInputBindings bindings() {
      return prepared.bindings();
    }
  }

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
    List<InspectionResult> inspections = bindings.session().inspections;
    List<AssertionResult> hostAssertions = bindings.session().hostAssertions;
    List<dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation> hostPreservation =
        bindings.session().preservation;
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
    WorkbookArtifactIo.@org.jspecify.annotations.Nullable MaterializedWorkbook materialized = null;
    var openContext =
        new dev.erst.gridgrind.core.problem.ProblemContext.OpenWorkbook(
            ExecutionRequestPaths.requestShape(request),
            ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory()),
            java.util.Optional.empty());
    try {
      materialized =
          call.prepared()
              .source()
              .takeMaterialized()
              .orElseThrow(
                  () -> new IllegalStateException("Admitted event source was not retained"));
    } catch (Exception exception) {
      return failedOpen(
          materialized,
          executionContext,
          openPhase,
          GridGrindProblems.fromException(exception, openContext));
    }
    if (!java.nio.file.Files.isRegularFile(
        materialized.workbookPath(), java.nio.file.LinkOption.NOFOLLOW_LINKS)) {
      return failedOpen(
          materialized,
          executionContext,
          openPhase,
          GridGrindProblems.problem(
              dev.erst.gridgrind.core.problem.GridGrindProblemCode.IO_ERROR,
              "Admitted workbook source materialization is unavailable",
              openContext,
              List.of()));
    }
    openPhase.succeed();
    dev.erst.gridgrind.core.grant.AcceptancePolicy hostAcceptancePolicy =
        bindings.acceptancePolicy();
    for (int stepIndex = 0; stepIndex < plan.steps().size(); stepIndex++) {
      var prepared =
          (dev.erst.gridgrind.core.plan.PreparedInspectionStep) plan.steps().get(stepIndex);
      InspectionStep inspectionStep = prepared.authoredStep();
      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, inspectionStep);
      try {
        inspections.add(stepSupport.executeEventInspection(materialized.workbookPath(), prepared));
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
              bindings.executionGrant().acceptancePolicy(),
              hostAcceptancePolicy,
              materialized.workbookPath(),
              ExecutionRequestPaths.workbookLocationFor(
                  request.source(), request.persistence(), bindings.workingDirectory()),
              hostAssertions);
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

  private WorkbookResult failedOpen(
      WorkbookArtifactIo.@org.jspecify.annotations.Nullable MaterializedWorkbook materialized,
      DirectEventReadContext context,
      ExecutionJournalRecorder.PhaseHandle phase,
      GridGrindProblemDetail.Problem problem) {
    phase.fail(problem.code());
    return responseSupport.closeReadableWorkbook(
        materialized,
        ExecutionResults.failureResponseWithoutPlanOutcomeEvent(context.failure(problem)),
        context.request(),
        context.journal(),
        problem.code(),
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
        dev.erst.gridgrind.core.grant.AcceptancePolicy authoredPolicy,
        dev.erst.gridgrind.core.grant.AcceptancePolicy policy,
        java.nio.file.Path workbookPath,
        dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation workbookLocation,
        List<AssertionResult> assertions)
        throws IOException, AssertionFailedException;
  }
}
