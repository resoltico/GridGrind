package dev.erst.gridgrind.core.validation;

import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Complete or partially bound facts supplied to the canonical static request contract. */
public record WorkbookStaticRequest(
    Optional<WorkbookPlan.WorkbookSource> source,
    Optional<WorkbookPlan.WorkbookPersistence> persistence,
    Optional<ExecutionPolicySpec> execution,
    List<WorkbookStaticStep> steps) {
  public WorkbookStaticRequest {
    source = copy(source, "source");
    persistence = copy(persistence, "persistence");
    execution = copy(execution, "execution");
    Objects.requireNonNull(steps, "steps must not be null");
    steps = List.copyOf(steps);
    for (int index = 0; index < steps.size(); index++) {
      WorkbookStaticStep step =
          Objects.requireNonNull(steps.get(index), "steps must not contain nulls");
      if (step.index() != index) {
        throw new IllegalArgumentException("steps must use contiguous authored indexes");
      }
    }
  }

  private static <T> Optional<T> copy(Optional<T> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    value.ifPresent(
        candidate -> Objects.requireNonNull(candidate, fieldName + " must not contain null"));
    return value;
  }
}
