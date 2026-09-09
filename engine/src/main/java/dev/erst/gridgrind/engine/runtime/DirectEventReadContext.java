package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
import org.jspecify.annotations.Nullable;

/** Request-scoped direct-event facts and evidence retained through readable-workbook closure. */
record DirectEventReadContext(
    GridGrindProtocolVersion protocolVersion,
    WorkbookPlan request,
    List<RequestWarning> warnings,
    ExecutionJournalRecorder journal,
    CalculationReport calculation,
    List<AssertionResult> hostAssertions,
    List<Preservation> preservation,
    List<InspectionResult> inspections) {
  ExecutionFailure failure(GridGrindProblemDetail.Problem problem) {
    return failure(problem, null, null);
  }

  ExecutionFailure failure(
      GridGrindProblemDetail.Problem problem, int failedStepIndex, String failedStepId) {
    return failure(problem, Integer.valueOf(failedStepIndex), failedStepId);
  }

  private ExecutionFailure failure(
      GridGrindProblemDetail.Problem problem,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return new ExecutionFailure(
        new ExecutionFailure.Context(protocolVersion, journal, request, calculation),
        new ExecutionFailure.Artifacts(
            warnings, List.of(), hostAssertions, preservation, inspections),
        new ExecutionFailure.Detail(problem, failedStepIndex, failedStepId));
  }
}
