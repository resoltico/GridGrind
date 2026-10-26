package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import java.util.List;
import java.util.Objects;

/** Shared workbook-plan skeleton helpers for shipped example requests. */
final class ExampleWorkbookPlans {
  private ExampleWorkbookPlans() {}

  static WorkbookPlan plan(
      PlanId planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      WorkbookStep... steps) {
    return WorkbookPlan.identified(
        planId,
        source,
        persistence,
        Objects.requireNonNull(execution, "execution must not be null"),
        dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
        List.of(steps));
  }

  static WorkbookPlan defaultExecutionPlan(
      PlanId planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookStep... steps) {
    return plan(planId, source, persistence, ExecutionPolicySpec.defaults(), steps);
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
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none());
  }
}
