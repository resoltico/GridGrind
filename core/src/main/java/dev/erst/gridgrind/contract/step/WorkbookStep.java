package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import tools.jackson.databind.annotation.JsonDeserialize;

/** One ordered executable step in a workbook plan. */
@JsonDeserialize(using = WorkbookStepJsonDeserializer.class)
public sealed interface WorkbookStep permits MutationStep, AssertionStep, InspectionStep {

  /** Stable caller-provided identifier unique within the containing plan. */
  String stepId();

  /** Semantic workbook target selected by this step. */
  Selector target();

  /** High-level step kind name used in diagnostics and discovery surfaces. */
  String stepKind();
}
