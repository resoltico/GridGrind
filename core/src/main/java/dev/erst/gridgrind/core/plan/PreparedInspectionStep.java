package dev.erst.gridgrind.core.plan;

import java.util.Objects;

/** An original inspection and its bound selector, kept distinct for diagnostic provenance. */
public record PreparedInspectionStep(InspectionStep authoredStep, Selector target)
    implements PreparedStep {
  public PreparedInspectionStep {
    Objects.requireNonNull(authoredStep, "authoredStep must not be null");
    Objects.requireNonNull(target, "target must not be null");
    SourceBindingRelationship.requireBinding(authoredStep.target(), target);
  }
}
