package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.ExecutionModeInput;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.util.ArrayList;
import java.util.List;
import org.jspecify.annotations.Nullable;

/** Mutable request-scoped evidence collected during one streaming-write workflow. */
record StreamingWorkflowContext(
    GridGrindProtocolVersion protocolVersion,
    WorkbookPlan request,
    ExecutionModeInput executionMode,
    List<RequestWarning> warnings,
    ExecutionJournalRecorder journal,
    WorkbookLocation workbookLocation,
    List<AssertionResult> planAssertions,
    List<AssertionResult> hostAssertions,
    List<Preservation> preservation,
    List<InspectionResult> inspections) {
  static StreamingWorkflowContext create(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    WorkbookLocation workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), bindings.workingDirectory());
    return new StreamingWorkflowContext(
        protocolVersion,
        request,
        executionMode,
        warnings,
        journal,
        workbookLocation,
        new ArrayList<>(),
        new ArrayList<>(),
        new ArrayList<>(),
        new ArrayList<>());
  }

  ExecutionFailure failure(CalculationReport calculation, GridGrindProblemDetail.Problem problem) {
    return failure(calculation, problem, null, null);
  }

  ExecutionFailure failure(
      CalculationReport calculation,
      GridGrindProblemDetail.Problem problem,
      int failedStepIndex,
      StepId failedStepId) {
    return failure(calculation, problem, Integer.valueOf(failedStepIndex), failedStepId);
  }

  private ExecutionFailure failure(
      CalculationReport calculation,
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable StepId failedStepId) {
    return new ExecutionFailure(
        new ExecutionFailure.Context(protocolVersion, journal, request, calculation),
        new ExecutionFailure.Artifacts(
            warnings, planAssertions, hostAssertions, preservation, inspections),
        new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId));
  }
}
