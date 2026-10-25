package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered executable step in a workbook plan. */
public sealed interface WorkbookStep permits MutationStep, AssertionStep, InspectionStep {

  /** Stable caller-provided identifier unique within the containing plan. */
  StepId stepId();

  /** Semantic workbook target selected by this step. */
  Selector target();

  /** High-level step kind name used in diagnostics and discovery surfaces. */
  String stepKind();
}
