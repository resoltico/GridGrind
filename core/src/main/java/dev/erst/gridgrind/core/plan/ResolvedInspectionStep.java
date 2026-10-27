package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Arrays;
import java.util.Objects;

/** One inspection with a source-free query and a target resolved against current workbook state. */
public record ResolvedInspectionStep(StepId stepId, ResolvedTarget target, InspectionQuery query) {
  public ResolvedInspectionStep {
    Objects.requireNonNull(stepId, "stepId must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(query, "query must not be null");
    if (Arrays.stream(InspectionQuery.allowedTargetTypes(query))
        .noneMatch(type -> type.isInstance(target))) {
      throw new IllegalArgumentException("Inspection target does not match " + query.queryType());
    }
  }
}
