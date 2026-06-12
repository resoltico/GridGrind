package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;

/** Shared step helpers for shipped example workbook plans. */
final class ExampleSteps {
  private ExampleSteps() {}

  static MutationStep step(String stepId, Selector target, MutationAction action) {
    return new MutationStep(stepId, target, action);
  }

  static InspectionStep read(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(stepId, target, query);
  }

  static AssertionStep assertStep(String stepId, Selector target, Assertion assertion) {
    return new AssertionStep(stepId, target, assertion);
  }
}
