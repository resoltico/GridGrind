package dev.erst.gridgrind.core.plan;

/** A source-materialized step whose workbook-dependent selector resolves at execution time. */
public sealed interface PreparedStep permits PreparedMutationStep, InspectionStep, AssertionStep {
  /** Authored identity and selector metadata used for diagnostics and journal events. */
  WorkbookStep authoredStep();
}
