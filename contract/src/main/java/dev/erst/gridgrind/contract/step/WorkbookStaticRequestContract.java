package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Canonical static request contract composed from record facts and operation contracts. */
public final class WorkbookStaticRequestContract {
  private WorkbookStaticRequestContract() {}

  /** Returns every independently provable static violation in authored request order. */
  public static List<WorkbookStaticViolation> validate(WorkbookStaticRequest request) {
    Objects.requireNonNull(request, "request must not be null");
    List<WorkbookStaticViolation> violations = new ArrayList<>();
    violations.addAll(WorkbookStaticTargetValidation.validate(request.steps()));
    violations.addAll(WorkbookStaticExecutionValidation.validate(request));
    violations.addAll(WorkbookStaticPersistenceValidation.validate(request));
    return List.copyOf(violations);
  }

  /** Adapts one complete plan into the same contract used for partial request analysis. */
  public static WorkbookStaticRequest from(WorkbookPlan plan) {
    Objects.requireNonNull(plan, "plan must not be null");
    return new WorkbookStaticRequest(
        Optional.of(plan.source()),
        Optional.of(plan.persistence()),
        Optional.of(plan.execution()),
        java.util.stream.IntStream.range(0, plan.steps().size())
            .mapToObj(index -> new WorkbookStaticStep(index, Optional.of(plan.steps().get(index))))
            .toList());
  }
}
