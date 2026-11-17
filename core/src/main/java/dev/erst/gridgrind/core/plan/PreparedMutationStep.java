package dev.erst.gridgrind.core.plan;

import java.util.Objects;

/** A mutation's original provenance and bound selector/action, with no second authored step. */
public final class PreparedMutationStep implements PreparedStep {
  private final MutationStep authoredStep;
  private final Selector target;
  private final PreparedMutationAction action;

  PreparedMutationStep(MutationStep authoredStep, MutationStep bound) {
    this.authoredStep = Objects.requireNonNull(authoredStep, "authoredStep must not be null");
    SourceBindingRelationship.requireBinding(authoredStep, bound);
    this.target = bound.target();
    this.action = MutationPreparation.prepare(bound.action());
  }

  @Override
  public MutationStep authoredStep() {
    return authoredStep;
  }

  /** Returns the selector with external key operands bound, before workbook-state resolution. */
  public Selector target() {
    return target;
  }

  /** Returns validated materialized operands; secret references remain unacquired. */
  public PreparedMutationAction action() {
    return action;
  }
}
