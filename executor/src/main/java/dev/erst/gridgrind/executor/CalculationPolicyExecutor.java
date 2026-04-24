package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityAssessment;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityKind;
import dev.erst.gridgrind.excel.ExcelFormulaCellTarget;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Applies the explicit calculation policy against a fully opened workbook. */
final class CalculationPolicyExecutor {
  private CalculationPolicyExecutor() {}

  static CalculationPolicyInput normalize(CalculationPolicyInput policy) {
    return policy == null ? new CalculationPolicyInput(null, false) : policy;
  }

  static boolean allowsEventRead(CalculationPolicyInput policy) {
    CalculationPolicyInput effective = normalize(policy);
    return effective.effectiveStrategy() instanceof CalculationStrategyInput.DoNotCalculate
        && !effective.markRecalculateOnOpen();
  }

  static boolean allowsStreamingWrite(CalculationPolicyInput policy) {
    CalculationPolicyInput effective = normalize(policy);
    return effective.effectiveStrategy() instanceof CalculationStrategyInput.DoNotCalculate;
  }

  static boolean requiresMutationPrefix(CalculationPolicyInput policy) {
    return switch (normalize(policy).effectiveStrategy()) {
      case CalculationStrategyInput.DoNotCalculate _ -> false;
      case CalculationStrategyInput.ClearCachesOnly _ -> true;
      case CalculationStrategyInput.EvaluateAll _ -> true;
      case CalculationStrategyInput.EvaluateTargets _ -> true;
    };
  }

  static CalculationReport notRequestedReport(CalculationPolicyInput policy) {
    CalculationPolicyInput effective = normalize(policy);
    return new CalculationReport(
        effective,
        null,
        new CalculationReport.Execution(
            CalculationExecutionStatus.NOT_REQUESTED, 0, false, false, null));
  }

  static PreflightOutcome preflight(ExcelWorkbook workbook, CalculationPolicyInput policy) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    CalculationPolicyInput effective = normalize(policy);
    return switch (effective.effectiveStrategy()) {
      case CalculationStrategyInput.DoNotCalculate _ -> new PreflightOutcome(null, 0, null);
      case CalculationStrategyInput.ClearCachesOnly _ -> new PreflightOutcome(null, 0, null);
      case CalculationStrategyInput.EvaluateAll _ -> preflightAll(workbook);
      case CalculationStrategyInput.EvaluateTargets evaluateTargets ->
          preflightTargets(workbook, evaluateTargets);
    };
  }

  static ExecutionOutcome execute(
      ExcelWorkbook workbook, CalculationPolicyInput policy, int evaluationTargetCount) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    CalculationPolicyInput effective = normalize(policy);
    return switch (effective.effectiveStrategy()) {
      case CalculationStrategyInput.DoNotCalculate _ -> executeDoNotCalculate(workbook, effective);
      case CalculationStrategyInput.ClearCachesOnly _ ->
          executeClearCachesOnly(workbook, effective);
      case CalculationStrategyInput.EvaluateAll _ ->
          executeEvaluateAll(workbook, effective, evaluationTargetCount);
      case CalculationStrategyInput.EvaluateTargets evaluateTargets ->
          executeEvaluateTargets(workbook, effective, evaluateTargets, evaluationTargetCount);
    };
  }

  static CalculationReport report(
      CalculationPolicyInput policy,
      CalculationReport.Preflight preflight,
      CalculationReport.Execution execution) {
    return new CalculationReport(normalize(policy), preflight, execution);
  }

  private static PreflightOutcome preflightAll(ExcelWorkbook workbook) {
    List<ExcelFormulaCapabilityAssessment> assessments =
        workbook.formulas().assessAllCapabilities();
    return buildPreflightOutcome(CalculationReport.Scope.WORKBOOK, assessments);
  }

  private static PreflightOutcome preflightTargets(
      ExcelWorkbook workbook, CalculationStrategyInput.EvaluateTargets strategy) {
    List<ExcelFormulaCellTarget> targets = toExcelFormulaTargets(strategy.cells());
    List<ExcelFormulaCapabilityAssessment> assessments =
        workbook.formulas().assessCapabilities(targets);
    return buildPreflightOutcome(CalculationReport.Scope.TARGETS, assessments);
  }

  private static PreflightOutcome buildPreflightOutcome(
      CalculationReport.Scope scope, List<ExcelFormulaCapabilityAssessment> assessments) {
    CalculationReport.Preflight report =
        new CalculationReport.Preflight(
            scope,
            assessments.size(),
            CalculationCapabilityMappings.summaryFor(assessments),
            toCapabilityReports(assessments));
    Optional<ExcelFormulaCapabilityAssessment> blocking = mostSevereBlockingAssessment(assessments);
    return new PreflightOutcome(
        report,
        assessments.size(),
        blocking
            .map(
                assessment ->
                    new FailureDetail(
                        CalculationCapabilityMappings.problemCodeFor(assessment).orElse(null),
                        Phase.PREFLIGHT,
                        assessment.sheetName(),
                        assessment.address(),
                        assessment.formula(),
                        Objects.requireNonNullElse(
                            assessment.message(),
                            "Calculation preflight found formulas that are not immediately evaluable."),
                        null))
            .orElse(null));
  }

  private static ExecutionOutcome executeDoNotCalculate(
      ExcelWorkbook workbook, CalculationPolicyInput policy) {
    boolean marked = false;
    if (policy.markRecalculateOnOpen()) {
      workbook.formulas().markRecalculateOnOpen();
      marked = true;
    }
    CalculationExecutionStatus status =
        marked ? CalculationExecutionStatus.SUCCEEDED : CalculationExecutionStatus.NOT_REQUESTED;
    return new ExecutionOutcome(
        new CalculationReport.Execution(status, 0, false, marked, null), null);
  }

  private static ExecutionOutcome executeClearCachesOnly(
      ExcelWorkbook workbook, CalculationPolicyInput policy) {
    try {
      workbook.formulas().clearCaches();
      boolean marked = false;
      if (policy.markRecalculateOnOpen()) {
        workbook.formulas().markRecalculateOnOpen();
        marked = true;
      }
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.SUCCEEDED, 0, true, marked, null),
          null);
    } catch (RuntimeException exception) {
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.FAILED,
              0,
              false,
              false,
              GridGrindProblems.messageFor(exception)),
          new FailureDetail(Phase.EXECUTION, exception));
    }
  }

  private static ExecutionOutcome executeEvaluateAll(
      ExcelWorkbook workbook, CalculationPolicyInput policy, int evaluationTargetCount) {
    try {
      workbook.formulas().evaluateAll();
      boolean marked = false;
      if (policy.markRecalculateOnOpen()) {
        workbook.formulas().markRecalculateOnOpen();
        marked = true;
      }
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.SUCCEEDED, evaluationTargetCount, false, marked, null),
          null);
    } catch (RuntimeException exception) {
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.FAILED,
              0,
              false,
              false,
              GridGrindProblems.messageFor(exception)),
          new FailureDetail(Phase.EXECUTION, exception));
    }
  }

  private static ExecutionOutcome executeEvaluateTargets(
      ExcelWorkbook workbook,
      CalculationPolicyInput policy,
      CalculationStrategyInput.EvaluateTargets strategy,
      int evaluationTargetCount) {
    List<ExcelFormulaCellTarget> targets = toExcelFormulaTargets(strategy.cells());
    try {
      workbook.formulas().evaluate(targets);
      boolean marked = false;
      if (policy.markRecalculateOnOpen()) {
        workbook.formulas().markRecalculateOnOpen();
        marked = true;
      }
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.SUCCEEDED, evaluationTargetCount, false, marked, null),
          null);
    } catch (RuntimeException exception) {
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.FAILED,
              0,
              false,
              false,
              GridGrindProblems.messageFor(exception)),
          new FailureDetail(Phase.EXECUTION, exception));
    }
  }

  private static List<ExcelFormulaCellTarget> toExcelFormulaTargets(
      List<CellSelector.QualifiedAddress> cells) {
    return cells.stream()
        .map(cell -> new ExcelFormulaCellTarget(cell.sheetName(), cell.address()))
        .toList();
  }

  private static List<CalculationReport.FormulaCapability> toCapabilityReports(
      List<ExcelFormulaCapabilityAssessment> assessments) {
    return assessments.stream()
        .map(
            assessment ->
                new CalculationReport.FormulaCapability(
                    new CellSelector.QualifiedAddress(assessment.sheetName(), assessment.address()),
                    assessment.formula(),
                    CalculationCapabilityMappings.capabilityKindFor(assessment.capability()),
                    CalculationCapabilityMappings.problemCodeFor(assessment).orElse(null),
                    assessment.message()))
        .toList();
  }

  private static Optional<ExcelFormulaCapabilityAssessment> mostSevereBlockingAssessment(
      List<ExcelFormulaCapabilityAssessment> assessments) {
    return assessments.stream()
        .filter(assessment -> assessment.capability() != ExcelFormulaCapabilityKind.EVALUABLE_NOW)
        .min(Comparator.comparingInt(CalculationCapabilityMappings::severityRank));
  }

  record PreflightOutcome(
      CalculationReport.Preflight report, int evaluationTargetCount, FailureDetail failure) {
    PreflightOutcome {
      if (evaluationTargetCount < 0) {
        throw new IllegalArgumentException("evaluationTargetCount must be >= 0");
      }
    }
  }

  record ExecutionOutcome(CalculationReport.Execution report, FailureDetail failure) {
    ExecutionOutcome {
      Objects.requireNonNull(report, "report must not be null");
    }
  }

  /** Distinguishes failures raised during preflight from failures raised during execution. */
  enum Phase {
    /** Failure raised while classifying formula capability before evaluation begins. */
    PREFLIGHT,
    /** Failure raised while executing the requested calculation strategy. */
    EXECUTION
  }

  record FailureDetail(
      GridGrindProblemCode code,
      Phase phase,
      String sheetName,
      String address,
      String formula,
      String message,
      RuntimeException exception) {
    FailureDetail {
      Objects.requireNonNull(phase, "phase must not be null");
      requireNonBlank(message, "message");
      if (code == null && exception == null) {
        throw new IllegalArgumentException("code or exception must be present");
      }
    }

    FailureDetail(Phase phase, RuntimeException exception) {
      this(null, phase, null, null, null, GridGrindProblems.messageFor(exception), exception);
    }
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }
}
