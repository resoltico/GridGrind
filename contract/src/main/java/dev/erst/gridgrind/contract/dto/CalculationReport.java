package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Structured response-side report for explicit formula calculation policy, preflight, and
 * execution.
 */
public record CalculationReport(
    CalculationPolicyInput policy,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Preflight> preflight,
    Execution execution) {
  public CalculationReport {
    Objects.requireNonNull(policy, "policy must not be null");
    preflight = Objects.requireNonNullElseGet(preflight, Optional::empty);
    Objects.requireNonNull(execution, "execution must not be null");
  }

  /** Creates one calculation report without a preflight section. */
  public CalculationReport(CalculationPolicyInput policy, Execution execution) {
    this(policy, Optional.empty(), execution);
  }

  /** Creates one calculation report with a concrete preflight section. */
  public CalculationReport(
      CalculationPolicyInput policy, Preflight preflight, Execution execution) {
    this(policy, Optional.of(preflight), execution);
  }

  /** Returns the default report emitted when calculation was not requested. */
  public static CalculationReport notRequested() {
    return new CalculationReport(
        CalculationPolicyInput.defaults(),
        new Execution(CalculationExecutionStatus.NOT_REQUESTED, 0, false, false, Optional.empty()));
  }

  @JsonCreator
  static CalculationReport create(
      @JsonProperty("policy") CalculationPolicyInput policy,
      @JsonProperty("preflight") Optional<Preflight> preflight,
      @JsonProperty("execution") Execution execution) {
    return new CalculationReport(
        policy == null ? CalculationPolicyInput.defaults() : policy, preflight, execution);
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
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<GridGrindProblemCode> problemCode,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> message) {
    public FormulaCapability {
      Objects.requireNonNull(cell, "cell must not be null");
      WorkbookPlan.requireNonBlank(formula, "formula");
      Objects.requireNonNull(capability, "capability must not be null");
      problemCode = Objects.requireNonNullElseGet(problemCode, Optional::empty);
      message = normalizeOptional(message, "message");
      if (capability == FormulaCapabilityKind.EVALUABLE_NOW && problemCode.isPresent()) {
        throw new IllegalArgumentException(
            "problemCode is not permitted for EVALUABLE_NOW formulas");
      }
      if (capability == FormulaCapabilityKind.EVALUABLE_NOW && message.isPresent()) {
        throw new IllegalArgumentException("message is not permitted for EVALUABLE_NOW formulas");
      }
      if (capability != FormulaCapabilityKind.EVALUABLE_NOW && problemCode.isEmpty()) {
        throw new IllegalArgumentException(
            "problemCode must be present for non-evaluable formula capabilities");
      }
    }
  }

  /** Concrete execution outcome for the authored calculation policy. */
  public record Execution(
      CalculationExecutionStatus status,
      int evaluatedFormulaCount,
      boolean cachesCleared,
      boolean markRecalculateOnOpenApplied,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> message) {
    public Execution {
      Objects.requireNonNull(status, "status must not be null");
      if (evaluatedFormulaCount < 0) {
        throw new IllegalArgumentException("evaluatedFormulaCount must be >= 0");
      }
      message = normalizeOptional(message, "message");
    }

    /** Creates one execution report without a message. */
    public Execution(
        CalculationExecutionStatus status,
        int evaluatedFormulaCount,
        boolean cachesCleared,
        boolean markRecalculateOnOpenApplied) {
      this(
          status,
          evaluatedFormulaCount,
          cachesCleared,
          markRecalculateOnOpenApplied,
          Optional.empty());
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

  private static Optional<String> normalizeOptional(Optional<String> value, String fieldName) {
    Optional<String> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    if (normalized.isPresent()) {
      normalized = Optional.of(WorkbookPlan.requireNonBlank(normalized.orElseThrow(), fieldName));
    }
    return normalized;
  }
}
