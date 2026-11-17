package dev.erst.gridgrind.core.plan;

import java.util.Objects;

/** An original assertion and its bound selector, kept distinct for diagnostic provenance. */
public record PreparedAssertionStep(AssertionStep authoredStep, Selector target)
    implements PreparedStep {
  public PreparedAssertionStep {
    Objects.requireNonNull(authoredStep, "authoredStep must not be null");
    Objects.requireNonNull(target, "target must not be null");
    SourceBindingRelationship.requireBinding(authoredStep.target(), target);
  }
}
