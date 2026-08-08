package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;
import java.util.Optional;

/** Bound root-level request components, separate from independently decoded steps. */
record RequestBoundRoot(
    Optional<GridGrindProtocolVersion> protocolVersion,
    Optional<String> planId,
    Optional<WorkbookPlan.WorkbookSource> source,
    Optional<WorkbookPlan.WorkbookPersistence> persistence,
    Optional<ExecutionPolicyInput> execution,
    Optional<FormulaEnvironmentInput> formulaEnvironment) {
  RequestBoundRoot {
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
