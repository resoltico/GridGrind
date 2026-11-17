package dev.erst.gridgrind.core.plan;

import java.util.List;
import java.util.Objects;

/**
 * Source-materialized execution steps plus request metadata retained for diagnostics and policy.
 */
public record ResolvedPlan(WorkbookPlan request, List<PreparedStep> steps) {
  public ResolvedPlan {
    Objects.requireNonNull(request, "request must not be null");
    steps = List.copyOf(steps);
    if (request.steps().size() != steps.size()) {
      throw new IllegalArgumentException("Prepared steps must retain request order and identity");
    }
    for (int index = 0; index < steps.size(); index++) {
      WorkbookStep original = request.steps().get(index);
      if (!original.equals(steps.get(index).authoredStep())) {
        throw new IllegalArgumentException(
            "Prepared steps must retain their complete authored provenance");
      }
    }
  }
}
