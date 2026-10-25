package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation;
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
