package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.util.List;
import java.util.Objects;

/**
 * Structured response-side report for explicit formula calculation policy, preflight, and
 * execution.
 */
public record CalculationReport(
    CalculationPolicyInput policy,
    @JsonInclude(JsonInclude.Include.NON_NULL) Preflight preflight,
    Execution execution) {
  public CalculationReport {
    policy = policy == null ? new CalculationPolicyInput(null, false) : policy;
    Objects.requireNonNull(execution, "execution must not be null");
  }

  /** Returns the default report emitted when calculation was not requested. */
  public static CalculationReport notRequested() {
    return new CalculationReport(
        new CalculationPolicyInput(null, false),
        null,
        new Execution(CalculationExecutionStatus.NOT_REQUESTED, 0, false, false, null));
  }

  /** Structured preflight report emitted before any server-side calculation attempt. */
  public record Preflight(
      Scope scope, int checkedFormulaCount, Summary summary, List<FormulaCapability> formulas) {
    public Preflight {
      Objects.requireNonNull(scope, "scope must not be null");
      if (checkedFormulaCount < 0) {
        throw new IllegalArgumentException("checkedFormulaCount must be >= 0");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      formulas = copyValues(formulas, "formulas");
      if (checkedFormulaCount != formulas.size()) {
        throw new IllegalArgumentException("checkedFormulaCount must equal formulas.size()");
      }
      if (summary.evaluableNowCount()
              + summary.unevaluableNowCount()
              + summary.unparseableByPoiCount()
          != checkedFormulaCount) {
        throw new IllegalArgumentException("summary counts must add up to checkedFormulaCount");
      }
    }
  }

  /** Workbook-wide or target-list preflight scope. */
  public enum Scope {
    WORKBOOK,
    TARGETS
  }

  /** Aggregate capability counts for one calculation preflight run. */
  public record Summary(int evaluableNowCount, int unevaluableNowCount, int unparseableByPoiCount) {
    public Summary {
      if (evaluableNowCount < 0 || unevaluableNowCount < 0 || unparseableByPoiCount < 0) {
        throw new IllegalArgumentException("summary counts must be >= 0");
      }
    }
  }

  /** One classified formula capability entry returned by calculation preflight. */
  public record FormulaCapability(
      CellSelector.QualifiedAddress cell,
      String formula,
      FormulaCapabilityKind capability,
      @JsonInclude(JsonInclude.Include.NON_NULL) GridGrindProblemCode problemCode,
      @JsonInclude(JsonInclude.Include.NON_NULL) String message) {
    public FormulaCapability {
      Objects.requireNonNull(cell, "cell must not be null");
      WorkbookPlan.requireNonBlank(formula, "formula");
      Objects.requireNonNull(capability, "capability must not be null");
      if (capability == FormulaCapabilityKind.EVALUABLE_NOW && problemCode != null) {
        throw new IllegalArgumentException(
            "problemCode is not permitted for EVALUABLE_NOW formulas");
      }
      if (capability != FormulaCapabilityKind.EVALUABLE_NOW && problemCode == null) {
        throw new IllegalArgumentException(
            "problemCode must be present for non-evaluable formula capabilities");
      }
      if (message != null) {
        WorkbookPlan.requireNonBlank(message, "message");
      }
    }
  }

  /** Concrete execution outcome for the authored calculation policy. */
  public record Execution(
      CalculationExecutionStatus status,
      int evaluatedFormulaCount,
      boolean cachesCleared,
      boolean markRecalculateOnOpenApplied,
      @JsonInclude(JsonInclude.Include.NON_NULL) String message) {
    public Execution {
      Objects.requireNonNull(status, "status must not be null");
      if (evaluatedFormulaCount < 0) {
        throw new IllegalArgumentException("evaluatedFormulaCount must be >= 0");
      }
      if (message != null) {
        WorkbookPlan.requireNonBlank(message, "message");
      }
    }
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new java.util.ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
