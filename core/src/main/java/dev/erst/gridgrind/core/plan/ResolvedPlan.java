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
    if (!request.steps().equals(steps.stream().map(PreparedStep::authoredStep).toList())) {
      throw new IllegalArgumentException("Prepared steps must retain request order and identity");
    }
  }
}
