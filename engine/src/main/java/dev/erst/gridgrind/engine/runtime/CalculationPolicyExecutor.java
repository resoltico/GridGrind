package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationExecutionStatus;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaCapabilityKind;
import dev.erst.gridgrind.core.model.formula.FormulaCellTarget;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityAssessment;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityKind;
import dev.erst.gridgrind.excel.ExcelFormulaCellTarget;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Applies the explicit calculation policy against a fully opened workbook. */
final class CalculationPolicyExecutor {
  private CalculationPolicyExecutor() {}

  static CalculationPolicySpec normalize(CalculationPolicySpec policy) {
    return policy == null ? CalculationPolicySpec.defaults() : policy;
  }

  static boolean allowsEventRead(CalculationPolicySpec policy) {
    return normalize(policy).allowsEventRead();
  }

  static boolean allowsStreamingWrite(CalculationPolicySpec policy) {
    return normalize(policy).allowsStreamingWrite();
  }

  static boolean requiresMutationPrefix(CalculationPolicySpec policy) {
    return normalize(policy).requiresMutationPrefix();
  }

  static CalculationReport notRequestedReport(CalculationPolicySpec policy) {
    CalculationPolicySpec effective = normalize(policy);
    return new CalculationReport(
        effective,
        new CalculationReport.Execution(CalculationExecutionStatus.NOT_REQUESTED, 0, false, false));
  }

  static PreflightOutcome preflight(ExcelWorkbook workbook, CalculationPolicySpec policy) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    CalculationPolicySpec effective = normalize(policy);
    return switch (effective.effectiveStrategy()) {
      case CalculationStrategySpec.DoNotCalculate _ ->
          new PreflightOutcome(Optional.empty(), 0, false, Optional.empty());
      case CalculationStrategySpec.DeferredCalculation _ -> preflightLenientAll(workbook);
      case CalculationStrategySpec.ClearCachesOnly _ ->
          new PreflightOutcome(Optional.empty(), 0, false, Optional.empty());
      case CalculationStrategySpec.EvaluateAll _ -> preflightLenientAll(workbook);
      case CalculationStrategySpec.EvaluateTargets evaluateTargets ->
          preflightLenientTargets(workbook, evaluateTargets);
      case CalculationStrategySpec.RequireEvaluation _ -> preflightAll(workbook);
    };
  }

  static ExecutionOutcome execute(
      ExcelWorkbook workbook,
      CalculationPolicySpec policy,
      int evaluationTargetCount,
      boolean hasUnevaluableFormula) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    CalculationPolicySpec effective = normalize(policy);
    return switch (effective.effectiveStrategy()) {
      case CalculationStrategySpec.DoNotCalculate _ -> executeDoNotCalculate(workbook, effective);
      case CalculationStrategySpec.DeferredCalculation _ -> executeDeferred(workbook, effective);
      case CalculationStrategySpec.ClearCachesOnly _ -> executeClearCachesOnly(workbook, effective);
      case CalculationStrategySpec.EvaluateAll _ ->
          hasUnevaluableFormula
              ? executePartial(workbook, effective)
              : executeEvaluateAll(workbook, effective, evaluationTargetCount);
      case CalculationStrategySpec.EvaluateTargets evaluateTargets ->
          hasUnevaluableFormula
              ? executePartial(workbook, effective)
              : executeEvaluateTargets(workbook, effective, evaluateTargets, evaluationTargetCount);
      case CalculationStrategySpec.RequireEvaluation _ ->
          executeEvaluateAll(workbook, effective, evaluationTargetCount);
    };
  }

  static CalculationReport report(
      CalculationPolicySpec policy,
      Optional<CalculationReport.Preflight> preflight,
      CalculationReport.Execution execution) {
    return new CalculationReport(normalize(policy), preflight, execution);
  }

  private static PreflightOutcome preflightAll(ExcelWorkbook workbook) {
    List<ExcelFormulaCapabilityAssessment> assessments =
        workbook.formulas().assessAllCapabilities();
    return buildPreflightOutcome(CalculationReport.Scope.WORKBOOK, assessments);
  }

  private static PreflightOutcome preflightLenientAll(ExcelWorkbook workbook) {
    return withoutBlockingFailure(preflightAll(workbook));
  }

  private static PreflightOutcome preflightTargets(
      ExcelWorkbook workbook, CalculationStrategySpec.EvaluateTargets strategy) {
    List<ExcelFormulaCellTarget> targets = toExcelFormulaTargets(strategy.cells());
    List<ExcelFormulaCapabilityAssessment> assessments =
        workbook.formulas().assessCapabilities(targets);
    return buildPreflightOutcome(CalculationReport.Scope.TARGETS, assessments);
  }

  private static PreflightOutcome preflightLenientTargets(
      ExcelWorkbook workbook, CalculationStrategySpec.EvaluateTargets strategy) {
    return withoutBlockingFailure(preflightTargets(workbook, strategy));
  }

  private static PreflightOutcome withoutBlockingFailure(PreflightOutcome outcome) {
    return new PreflightOutcome(
        outcome.report(),
        outcome.evaluationTargetCount(),
        outcome.hasUnevaluableFormula(),
        Optional.empty());
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
        Optional.of(report),
        assessments.size(),
        blocking.isPresent(),
        blocking.map(
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
                    null)));
  }

  private static ExecutionOutcome executeDoNotCalculate(
      ExcelWorkbook workbook, CalculationPolicySpec policy) {
    boolean marked = false;
    if (policy.markRecalculateOnOpen()) {
      workbook.formulas().markRecalculateOnOpen();
      marked = true;
    }
    CalculationExecutionStatus status =
        marked ? CalculationExecutionStatus.SUCCEEDED : CalculationExecutionStatus.NOT_REQUESTED;
    return new ExecutionOutcome(
        new CalculationReport.Execution(status, 0, false, marked), Optional.empty());
  }

  private static ExecutionOutcome executeDeferred(
      ExcelWorkbook workbook, CalculationPolicySpec policy) {
    boolean marked = false;
    if (policy.markRecalculateOnOpen()) {
      workbook.formulas().markRecalculateOnOpen();
      marked = true;
    }
    CalculationExecutionStatus status =
        marked ? CalculationExecutionStatus.SUCCEEDED : CalculationExecutionStatus.NOT_REQUESTED;
    return new ExecutionOutcome(
        new CalculationReport.Execution(status, 0, false, marked), Optional.empty());
  }

  private static ExecutionOutcome executeClearCachesOnly(
      ExcelWorkbook workbook, CalculationPolicySpec policy) {
    try {
      workbook.formulas().clearCaches();
      boolean marked = false;
      if (policy.markRecalculateOnOpen()) {
        workbook.formulas().markRecalculateOnOpen();
        marked = true;
      }
      return new ExecutionOutcome(
          new CalculationReport.Execution(CalculationExecutionStatus.SUCCEEDED, 0, true, marked),
          Optional.empty());
    } catch (RuntimeException exception) {
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.FAILED,
              0,
              false,
              false,
              Optional.of(GridGrindProblems.messageFor(exception))),
          Optional.of(new FailureDetail(Phase.EXECUTION, exception)));
    }
  }

  private static ExecutionOutcome executeEvaluateAll(
      ExcelWorkbook workbook, CalculationPolicySpec policy, int evaluationTargetCount) {
    try {
      workbook.formulas().evaluateAll();
      boolean marked = false;
      if (policy.markRecalculateOnOpen()) {
        workbook.formulas().markRecalculateOnOpen();
        marked = true;
      }
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.SUCCEEDED, evaluationTargetCount, false, marked),
          Optional.empty());
    } catch (RuntimeException exception) {
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.FAILED,
              0,
              false,
              false,
              Optional.of(GridGrindProblems.messageFor(exception))),
          Optional.of(new FailureDetail(Phase.EXECUTION, exception)));
    }
  }

  private static ExecutionOutcome executeEvaluateTargets(
      ExcelWorkbook workbook,
      CalculationPolicySpec policy,
      CalculationStrategySpec.EvaluateTargets strategy,
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
              CalculationExecutionStatus.SUCCEEDED, evaluationTargetCount, false, marked),
          Optional.empty());
    } catch (RuntimeException exception) {
      return new ExecutionOutcome(
          new CalculationReport.Execution(
              CalculationExecutionStatus.FAILED,
              0,
              false,
              false,
              Optional.of(GridGrindProblems.messageFor(exception))),
          Optional.of(new FailureDetail(Phase.EXECUTION, exception)));
    }
  }

  private static ExecutionOutcome executePartial(
      ExcelWorkbook workbook, CalculationPolicySpec policy) {
    boolean marked = false;
    if (policy.markRecalculateOnOpen()) {
      workbook.formulas().markRecalculateOnOpen();
      marked = true;
    }
    return new ExecutionOutcome(
        new CalculationReport.Execution(CalculationExecutionStatus.PARTIAL, 0, false, marked),
        Optional.empty());
  }

  private static List<ExcelFormulaCellTarget> toExcelFormulaTargets(List<FormulaCellTarget> cells) {
    return cells.stream()
        .map(cell -> new ExcelFormulaCellTarget(cell.sheetName().value(), cell.address().value()))
        .toList();
  }

  private static List<CalculationReport.FormulaCapability> toCapabilityReports(
      List<ExcelFormulaCapabilityAssessment> assessments) {
    return assessments.stream()
        .map(
            assessment -> {
              FormulaCapabilityKind capability =
                  CalculationCapabilityMappings.capabilityKindFor(assessment.capability());
              Optional<String> message =
                  capability == FormulaCapabilityKind.EVALUABLE_NOW
                      ? Optional.empty()
                      : Optional.ofNullable(assessment.message());
              return new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(
                      new SheetName(assessment.sheetName()), new CellAddress(assessment.address())),
                  assessment.formula(),
                  capability,
                  CalculationCapabilityMappings.problemCodeFor(assessment),
                  message);
            })
        .toList();
  }

  private static Optional<ExcelFormulaCapabilityAssessment> mostSevereBlockingAssessment(
      List<ExcelFormulaCapabilityAssessment> assessments) {
    return assessments.stream()
        .filter(assessment -> assessment.capability() != ExcelFormulaCapabilityKind.EVALUABLE_NOW)
        .min(Comparator.comparingInt(CalculationCapabilityMappings::severityRank));
  }

  record PreflightOutcome(
      Optional<CalculationReport.Preflight> report,
      int evaluationTargetCount,
      boolean hasUnevaluableFormula,
      Optional<FailureDetail> failure) {
    PreflightOutcome {
      report = Objects.requireNonNullElseGet(report, Optional::empty);
      failure = Objects.requireNonNullElseGet(failure, Optional::empty);
      if (evaluationTargetCount < 0) {
        throw new IllegalArgumentException("evaluationTargetCount must be >= 0");
      }
    }
  }

  record ExecutionOutcome(CalculationReport.Execution report, Optional<FailureDetail> failure) {
    ExecutionOutcome {
      Objects.requireNonNull(report, "report must not be null");
      failure = Objects.requireNonNullElseGet(failure, Optional::empty);
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
      @Nullable GridGrindProblemCode code,
      Phase phase,
      @Nullable String sheetName,
      @Nullable String address,
      @Nullable String formula,
      String message,
      @Nullable RuntimeException exception) {
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
