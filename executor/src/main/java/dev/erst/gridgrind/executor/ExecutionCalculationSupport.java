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
import java.util.Optional;

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
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), Optional.empty());
    }

    Optional<CalculationReport.Preflight> preflightReport = Optional.empty();
    int evaluationTargetCount = 0;
    if (!(effectivePolicy.effectiveStrategy() instanceof CalculationStrategyInput.DoNotCalculate)
        && !(effectivePolicy.effectiveStrategy()
            instanceof CalculationStrategyInput.ClearCachesOnly)) {
      ExecutionJournalRecorder.PhaseHandle preflightPhase = journal.beginCalculationPreflight();
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, effectivePolicy);
      preflightReport = preflight.report();
      evaluationTargetCount = preflight.evaluationTargetCount();
      if (preflight.failure().isPresent()) {
        CalculationPolicyExecutor.FailureDetail failure = preflight.failure().orElseThrow();
        GridGrindResponse.Problem problem = calculationProblemFor(request, failure);
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
                    Optional.of(failure.message()))),
            Optional.of(problem));
      }
      preflightPhase.succeed();
    }

    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    CalculationPolicyExecutor.ExecutionOutcome execution =
        CalculationPolicyExecutor.execute(workbook, effectivePolicy, evaluationTargetCount);
    CalculationReport report =
        CalculationPolicyExecutor.report(effectivePolicy, preflightReport, execution.report());
    if (execution.failure().isPresent()) {
      GridGrindResponse.Problem problem =
          calculationProblemFor(request, execution.failure().orElseThrow());
      executionPhase.fail("failed (" + problem.code() + ")");
      return new CalculationExecutionOutcome(report, Optional.of(problem));
    }
    executionPhase.succeed();
    return new CalculationExecutionOutcome(report, Optional.empty());
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
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), Optional.empty());
    }
    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    try {
      streamingCalculationApplier.apply(writer);
      executionPhase.succeed();
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.report(
              effectivePolicy,
              Optional.empty(),
              new CalculationReport.Execution(
                  CalculationExecutionStatus.SUCCEEDED, 0, false, true)),
          Optional.empty());
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          GridGrindProblems.fromException(
              exception,
              calculationContextFor(request, CalculationPolicyExecutor.Phase.EXECUTION, null));
      executionPhase.fail("failed (" + problem.code() + ")");
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.report(
              effectivePolicy,
              Optional.empty(),
              new CalculationReport.Execution(
                  CalculationExecutionStatus.FAILED,
                  0,
                  false,
                  false,
                  Optional.of(GridGrindProblems.messageFor(exception)))),
          Optional.of(problem));
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

  dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation calculationContextFor(
      WorkbookPlan request,
      CalculationPolicyExecutor.Phase phase,
      CalculationPolicyExecutor.FailureDetail failure) {
    return switch (phase) {
      case PREFLIGHT ->
          new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation.Preflight(
              ExecutionRequestPaths.requestShape(request), failureLocation(failure));
      case EXECUTION ->
          new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation.Execution(
              ExecutionRequestPaths.requestShape(request), failureLocation(failure));
    };
  }

  private static dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation failureLocation(
      CalculationPolicyExecutor.FailureDetail failure) {
    if (failure == null
        || failure.sheetName() == null
        || failure.address() == null
        || failure.formula() == null) {
      return dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.unknown();
    }
    return dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.formulaCell(
        failure.sheetName(), failure.address(), failure.formula());
  }

  record CalculationExecutionOutcome(
      CalculationReport report, Optional<GridGrindResponse.Problem> failure) {
    CalculationExecutionOutcome {
      Objects.requireNonNull(report, "report must not be null");
      failure = Objects.requireNonNullElseGet(failure, Optional::empty);
    }
  }
}
