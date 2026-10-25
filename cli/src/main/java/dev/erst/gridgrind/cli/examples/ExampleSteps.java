package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.plan.Assertion;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** Shared step helpers for shipped example workbook plans. */
final class ExampleSteps {
  private ExampleSteps() {}

  static MutationStep step(StepId stepId, Selector target, MutationAction action) {
    return new MutationStep(stepId, target, action);
  }

  static InspectionStep read(StepId stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(stepId, target, query);
  }

  static AssertionStep assertStep(StepId stepId, Selector target, Assertion assertion) {
    return new AssertionStep(stepId, target, assertion);
  }
}
