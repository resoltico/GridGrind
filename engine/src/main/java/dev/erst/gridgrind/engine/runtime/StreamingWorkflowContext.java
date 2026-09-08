package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
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
        new ExecutionFailure.Artifacts(
            warnings, planAssertions, hostAssertions, preservation, inspections),
        new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId));
  }
}
