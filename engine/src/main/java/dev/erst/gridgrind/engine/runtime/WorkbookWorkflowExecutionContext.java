package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.List;
import org.jspecify.annotations.Nullable;

/** Mutable-artifact context held while the full-XSSF workflow executes one request. */
record WorkbookWorkflowExecutionContext(
    ProtocolVersion protocolVersion,
    WorkbookPlan request,
    ExcelWorkbook workbook,
    ExecutionJournalRecorder journal,
    List<RequestWarning> warnings,
    List<AssertionResult> planAssertions,
    List<AssertionResult> hostAssertions,
    List<dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation> preservation,
    List<InspectionResult> inspections) {
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

  ExecutionFailure failure(
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
