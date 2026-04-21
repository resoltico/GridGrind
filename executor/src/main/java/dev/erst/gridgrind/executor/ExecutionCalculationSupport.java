package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.Objects;

/** Calculation-policy execution and problem shaping for request workflows. */
final class ExecutionCalculationSupport {
  private final StreamingCalculationApplier streamingCalculationApplier;

  ExecutionCalculationSupport(StreamingCalculationApplier streamingCalculationApplier) {
    this.streamingCalculationApplier =
        Objects.requireNonNull(
            streamingCalculationApplier, "streamingCalculationApplier must not be null");
  }

  CalculationExecutionOutcome executeCalculationPolicy(
      ExcelWorkbook workbook, WorkbookPlan request, ExecutionJournalRecorder journal) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    CalculationPolicyInput policy = request.calculationPolicy();
    CalculationPolicyInput effectivePolicy = CalculationPolicyExecutor.normalize(policy);
    if (effectivePolicy.isDefault()) {
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), null);
    }

    CalculationReport.Preflight preflightReport = null;
    int evaluationTargetCount = 0;
    if (!(effectivePolicy.effectiveStrategy() instanceof CalculationStrategyInput.DoNotCalculate)
        && !(effectivePolicy.effectiveStrategy()
            instanceof CalculationStrategyInput.ClearCachesOnly)) {
      ExecutionJournalRecorder.PhaseHandle preflightPhase = journal.beginCalculationPreflight();
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, effectivePolicy);
      preflightReport = preflight.report();
      evaluationTargetCount = preflight.evaluationTargetCount();
      if (preflight.failure() != null) {
        GridGrindResponse.Problem problem = calculationProblemFor(request, preflight.failure());
        preflightPhase.fail("failed (" + problem.code() + ")");
        return new CalculationExecutionOutcome(
            CalculationPolicyExecutor.report(
                effectivePolicy,
                preflightReport,
                new CalculationReport.Execution(
                    CalculationExecutionStatus.FAILED,
                    0,
                    false,
                    false,
                    preflight.failure().message())),
            problem);
      }
      preflightPhase.succeed();
    }

    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    CalculationPolicyExecutor.ExecutionOutcome execution =
        CalculationPolicyExecutor.execute(workbook, effectivePolicy, evaluationTargetCount);
    CalculationReport report =
        CalculationPolicyExecutor.report(effectivePolicy, preflightReport, execution.report());
    if (execution.failure() != null) {
      GridGrindResponse.Problem problem = calculationProblemFor(request, execution.failure());
      executionPhase.fail("failed (" + problem.code() + ")");
      return new CalculationExecutionOutcome(report, problem);
    }
    executionPhase.succeed();
    return new CalculationExecutionOutcome(report, null);
  }

  CalculationExecutionOutcome executeStreamingCalculationPolicy(
      ExcelStreamingWorkbookWriter writer, WorkbookPlan request, ExecutionJournalRecorder journal) {
    Objects.requireNonNull(writer, "writer must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    CalculationPolicyInput policy = request.calculationPolicy();
    CalculationPolicyInput effectivePolicy = CalculationPolicyExecutor.normalize(policy);
    if (effectivePolicy.isDefault()) {
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), null);
    }
    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    try {
      streamingCalculationApplier.apply(writer);
      executionPhase.succeed();
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.report(
              effectivePolicy,
              null,
              new CalculationReport.Execution(
                  CalculationExecutionStatus.SUCCEEDED, 0, false, true, null)),
          null);
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          GridGrindProblems.fromException(
              exception,
              calculationContextFor(request, CalculationPolicyExecutor.Phase.EXECUTION, null));
      executionPhase.fail("failed (" + problem.code() + ")");
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.report(
              effectivePolicy,
              null,
              new CalculationReport.Execution(
                  CalculationExecutionStatus.FAILED,
                  0,
                  false,
                  false,
                  GridGrindProblems.messageFor(exception))),
          problem);
    }
  }

  GridGrindResponse.Problem calculationProblemFor(
      WorkbookPlan request, CalculationPolicyExecutor.FailureDetail failure) {
    if (failure.exception() != null) {
      return GridGrindProblems.fromException(
          failure.exception(), calculationContextFor(request, failure.phase(), failure));
    }
    return GridGrindProblems.problem(
        Objects.requireNonNull(failure.code(), "failure.code must not be null"),
        failure.message(),
        calculationContextFor(request, failure.phase(), failure),
        (Throwable) null);
  }

  GridGrindResponse.ProblemContext.ExecuteCalculation calculationContextFor(
      WorkbookPlan request,
      CalculationPolicyExecutor.Phase phase,
      CalculationPolicyExecutor.FailureDetail failure) {
    return switch (phase) {
      case PREFLIGHT ->
          new GridGrindResponse.ProblemContext.ExecuteCalculation.Preflight(
              ExecutionRequestPaths.reqSourceType(request),
              ExecutionRequestPaths.reqPersistenceType(request),
              failure == null ? null : failure.sheetName(),
              failure == null ? null : failure.address(),
              failure == null ? null : failure.formula());
      case EXECUTION ->
          new GridGrindResponse.ProblemContext.ExecuteCalculation.Execution(
              ExecutionRequestPaths.reqSourceType(request),
              ExecutionRequestPaths.reqPersistenceType(request),
              failure == null ? null : failure.sheetName(),
              failure == null ? null : failure.address(),
              failure == null ? null : failure.formula());
    };
  }

  record CalculationExecutionOutcome(CalculationReport report, GridGrindResponse.Problem failure) {
    CalculationExecutionOutcome {
      Objects.requireNonNull(report, "report must not be null");
    }
  }
}
