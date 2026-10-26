package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Independently bound request fragments retained when another branch is structurally malformed. */
public final class RequestBoundFragments {
  private final RequestBoundRoot root;
  private final Optional<List<Step>> steps;

  /** Retains independently bound root fields and step fragments after intake. */
  public RequestBoundFragments(RequestBoundRoot root, Optional<List<Step>> steps) {
    this.root = Objects.requireNonNull(root, "root must not be null");
    this.steps = copySteps(steps);
  }

  /** Returns the independently bound protocol version when its branch has no structural defect. */
  public Optional<GridGrindProtocolVersion> protocolVersion() {
    return root.protocolVersion();
  }

  /** Returns the independently bound optional plan identifier. */
  public Optional<PlanId> planId() {
    return root.planId();
  }

  /** Returns the independently bound source branch when it is structurally valid. */
  public Optional<WorkbookPlan.WorkbookSource> source() {
    return root.source();
  }

  /** Returns the independently bound persistence branch when it is structurally valid. */
  public Optional<WorkbookPlan.WorkbookPersistence> persistence() {
    return root.persistence();
  }

  /** Returns the explicitly bound or omission-defaulted execution policy. */
  public Optional<ExecutionPolicySpec> execution() {
    return root.execution();
  }

  /** Returns the explicitly bound or omission-defaulted formula environment. */
  public Optional<FormulaEnvironmentSpec> formulaEnvironment() {
    return root.formulaEnvironment();
  }

  /**
   * Returns every independently bound step only when the outer {@code steps} field is unambiguous.
   */
  public Optional<List<Step>> steps() {
    return steps;
  }

  /** Returns a fully bound plan only when every top-level component and every step bound. */
  public Optional<WorkbookPlan> completePlan() {
    if (root.protocolVersion().isEmpty()
        || root.source().isEmpty()
        || root.persistence().isEmpty()
        || root.execution().isEmpty()
        || root.formulaEnvironment().isEmpty()
        || steps.isEmpty()
        || steps.orElseThrow().stream().anyMatch(step -> step.value().isEmpty())) {
      return Optional.empty();
    }
    return Optional.of(
        new WorkbookPlan(
            root.protocolVersion().orElseThrow(),
            root.planId(),
            root.source().orElseThrow(),
            root.persistence().orElseThrow(),
            root.execution().orElseThrow(),
            root.formulaEnvironment().orElseThrow(),
            steps.orElseThrow().stream().map(step -> step.value().orElseThrow()).toList()));
  }

  /** One independently decoded step fragment at its authored array index. */
  public record Step(int index, Optional<WorkbookStep> value) {
    public Step {
      if (index < 0) {
        throw new IllegalArgumentException("index must not be negative");
      }
      value = copy(value, "value");
    }
  }

  private static <T> Optional<T> copy(Optional<T> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    value.ifPresent(
        candidate -> Objects.requireNonNull(candidate, fieldName + " must not contain null"));
    return value;
  }

  private static Optional<List<Step>> copySteps(Optional<List<Step>> steps) {
    Objects.requireNonNull(steps, "steps must not be null");
    return steps.map(
        values -> {
          List<Step> copy = List.copyOf(Objects.requireNonNull(values, "steps must not be null"));
          copy.forEach(step -> Objects.requireNonNull(step, "steps must not contain null"));
          return copy;
        });
  }
}
