package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindWarningCode;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.RequestWarningLocation;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

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
      journal.markCalculationPreflightNotRequested();
      journal.markCalculationExecutionNotRequested();
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), Optional.empty());
    }

    Optional<CalculationReport.Preflight> preflightReport = Optional.empty();
    int evaluationTargetCount = 0;
    boolean hasUnevaluableFormula = false;
    if (!(effectivePolicy.effectiveStrategy() instanceof CalculationStrategyInput.DoNotCalculate)
        && !(effectivePolicy.effectiveStrategy()
            instanceof CalculationStrategyInput.ClearCachesOnly)) {
      ExecutionJournalRecorder.PhaseHandle preflightPhase = journal.beginCalculationPreflight();
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, effectivePolicy);
      preflightReport = preflight.report();
      evaluationTargetCount = preflight.evaluationTargetCount();
      hasUnevaluableFormula = preflight.hasUnevaluableFormula();
      if (preflight.failure().isPresent()) {
        CalculationPolicyExecutor.FailureDetail failure = preflight.failure().orElseThrow();
        GridGrindProblemDetail.Problem problem = calculationProblemFor(request, failure);
        preflightPhase.fail(problem.code());
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
    } else {
      journal.markCalculationPreflightNotRequested();
    }

    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    CalculationPolicyExecutor.ExecutionOutcome execution =
        CalculationPolicyExecutor.execute(
            workbook, effectivePolicy, evaluationTargetCount, hasUnevaluableFormula);
    CalculationReport report =
        CalculationPolicyExecutor.report(effectivePolicy, preflightReport, execution.report());
    if (execution.failure().isPresent()) {
      GridGrindProblemDetail.Problem problem =
          calculationProblemFor(request, execution.failure().orElseThrow());
      executionPhase.fail(problem.code());
      return new CalculationExecutionOutcome(report, Optional.of(problem));
    }
    executionPhase.succeed();
    return new CalculationExecutionOutcome(
        report, Optional.empty(), formulaWarnings(preflightReport));
  }

  CalculationExecutionOutcome executeStreamingCalculationPolicy(
      ExcelStreamingWorkbookWriter writer, WorkbookPlan request, ExecutionJournalRecorder journal) {
    Objects.requireNonNull(writer, "writer must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    CalculationPolicyInput policy = request.calculationPolicy();
    CalculationPolicyInput effectivePolicy = CalculationPolicyExecutor.normalize(policy);
    if (effectivePolicy.isDefault()) {
      journal.markCalculationPreflightNotRequested();
      journal.markCalculationExecutionNotRequested();
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), Optional.empty());
    }
    journal.markCalculationPreflightNotRequested();
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
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              calculationContextFor(request, CalculationPolicyExecutor.Phase.EXECUTION, null));
      executionPhase.fail(problem.code());
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

  GridGrindProblemDetail.Problem calculationProblemFor(
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

  private static List<RequestWarning> formulaWarnings(
      Optional<CalculationReport.Preflight> preflightReport) {
    return preflightReport.stream()
        .flatMap(report -> report.formulas().stream())
        .filter(
            capability ->
                capability.capability()
                    != dev.erst.gridgrind.contract.dto.FormulaCapabilityKind.EVALUABLE_NOW)
        .map(
            capability ->
                new RequestWarning(
                    GridGrindWarningCode.FORMULA_NOT_EVALUATED,
                    new RequestWarningLocation.FormulaCell(
                        capability.cell().sheetName(),
                        capability.cell().address(),
                        capability.formula()),
                    capability
                        .message()
                        .orElse("Formula was not evaluated by the configured strategy.")))
        .toList();
  }

  dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation calculationContextFor(
      WorkbookPlan request,
      CalculationPolicyExecutor.Phase phase,
      CalculationPolicyExecutor.@Nullable FailureDetail failure) {
    return switch (phase) {
      case PREFLIGHT ->
          new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation.Preflight(
              ExecutionRequestPaths.requestShape(request), failureLocation(failure));
      case EXECUTION ->
          new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation.Execution(
              ExecutionRequestPaths.requestShape(request), failureLocation(failure));
    };
  }

  private static dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
      failureLocation(CalculationPolicyExecutor.@Nullable FailureDetail failure) {
    if (failure == null
        || failure.sheetName() == null
        || failure.address() == null
        || failure.formula() == null) {
      return dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
          .unknown();
    }
    return dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
        .formulaCell(failure.sheetName(), failure.address(), failure.formula());
  }

  record CalculationExecutionOutcome(
      CalculationReport report,
      Optional<GridGrindProblemDetail.Problem> failure,
      List<RequestWarning> warnings) {
    CalculationExecutionOutcome {
      Objects.requireNonNull(report, "report must not be null");
      failure = Objects.requireNonNullElseGet(failure, Optional::empty);
      warnings = List.copyOf(Objects.requireNonNullElseGet(warnings, List::of));
    }

    CalculationExecutionOutcome(
        CalculationReport report, Optional<GridGrindProblemDetail.Problem> failure) {
      this(report, failure, List.of());
    }
  }
}
