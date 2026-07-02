package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.List;
import java.util.Objects;

/** Shared workbook-plan skeleton helpers for shipped example requests. */
final class ExampleWorkbookPlans {
  private ExampleWorkbookPlans() {}

  static WorkbookPlan plan(
      String planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      WorkbookStep... steps) {
    return WorkbookPlan.identified(
        planId,
        source,
        persistence,
        Objects.requireNonNull(execution, "execution must not be null"),
        dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
        List.of(steps));
  }

  static WorkbookPlan defaultExecutionPlan(
      String planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookStep... steps) {
    return plan(planId, source, persistence, ExecutionPolicyInput.defaults(), steps);
  }

  static WorkbookPlan.WorkbookPersistence.SaveAs saveAs(String path) {
    return new WorkbookPlan.WorkbookPersistence.SaveAs(
        path, WorkbookPlan.WorkbookPersistence.IfExists.REPLACE);
  }
}
