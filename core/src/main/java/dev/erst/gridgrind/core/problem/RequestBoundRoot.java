package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import java.util.Objects;
import java.util.Optional;

/** Bound root-level request components, separate from independently decoded steps. */
public record RequestBoundRoot(
    Optional<GridGrindProtocolVersion> protocolVersion,
    Optional<PlanId> planId,
    Optional<WorkbookPlan.WorkbookSource> source,
    Optional<WorkbookPlan.WorkbookPersistence> persistence,
    Optional<ExecutionPolicySpec> execution,
    Optional<FormulaEnvironmentSpec> formulaEnvironment) {
  public RequestBoundRoot {
    protocolVersion = copy(protocolVersion, "protocolVersion");
    planId = copy(planId, "planId");
    source = copy(source, "source");
    persistence = copy(persistence, "persistence");
    execution = copy(execution, "execution");
    formulaEnvironment = copy(formulaEnvironment, "formulaEnvironment");
  }

  private static <T> Optional<T> copy(Optional<T> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    value.ifPresent(
        candidate -> Objects.requireNonNull(candidate, fieldName + " must not contain null"));
    return value;
  }
}
