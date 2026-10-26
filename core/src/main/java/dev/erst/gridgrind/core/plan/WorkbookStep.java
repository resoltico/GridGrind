package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.vocabulary.StepId;

/** One ordered executable step in a workbook plan. */
@dev.erst.gridgrind.core.metadata.Union(discriminator = "")
public sealed interface WorkbookStep
    permits dev.erst.gridgrind.core.plan.MutationStep,
        dev.erst.gridgrind.core.plan.AssertionStep,
        dev.erst.gridgrind.core.plan.InspectionStep {

  /** Stable caller-provided identifier unique within the containing plan. */
  StepId stepId();

  /** Semantic workbook target selected by this step. */
  Selector target();

  /** High-level step kind name used in diagnostics and discovery surfaces. */
  String stepKind();
}
