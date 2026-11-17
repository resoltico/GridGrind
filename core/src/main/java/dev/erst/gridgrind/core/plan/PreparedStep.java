package dev.erst.gridgrind.core.plan;

import java.util.Objects;

/** A step retaining original provenance separately from materialized execution operands. */
public sealed interface PreparedStep
    permits PreparedMutationStep, PreparedInspectionStep, PreparedAssertionStep {
  /** Original authored step used for diagnostics and journal events. */
  WorkbookStep authoredStep();

  /** Prepares a step whose inputs are already inline, without changing its provenance. */
  static PreparedStep prepare(WorkbookStep authored) {
    return prepare(authored, authored);
  }

  /**
   * Validates source binding and derives prepared operands from the corresponding authored step.
   */
  static PreparedStep prepare(WorkbookStep authored, WorkbookStep bound) {
    Objects.requireNonNull(authored, "authored must not be null");
    Objects.requireNonNull(bound, "bound must not be null");
    SourceBindingRelationship.requireBinding(authored, bound);
    return switch (authored) {
      case MutationStep mutation -> new PreparedMutationStep(mutation, (MutationStep) bound);
      case AssertionStep assertion -> new PreparedAssertionStep(assertion, bound.target());
      case InspectionStep inspection -> new PreparedInspectionStep(inspection, bound.target());
    };
  }
}
