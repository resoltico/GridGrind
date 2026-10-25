package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
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
        dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput.empty(),
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

  /** Returns a plain and unsigned SAVE_AS policy for one existing workbook source. */
  static WorkbookPlan.WorkbookPersistence.SaveAs saveAsExisting(String path) {
    return new WorkbookPlan.WorkbookPersistence.SaveAs(
        path,
        WorkbookPlan.WorkbookPersistence.IfExists.REPLACE,
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecurityInput.none());
  }
}
