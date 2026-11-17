package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.Objects;

/** Public request-shape queries needed by transport adapters before execution starts. */
public final class GridGrindRequestRequirements {
  private GridGrindRequestRequirements() {}

  /** Returns true when host acceptance operands require the bound standard-input payload. */
  public static boolean requiresStandardInput(dev.erst.gridgrind.core.grant.ExecutionGrant grant) {
    return dev.erst.gridgrind.engine.runtime.SourceBackedInputRequirements.requiresStandardInput(
        Objects.requireNonNull(grant).acceptancePolicy());
  }

  /** Returns true when any authored input in the request requires bound stdin bytes. */
  public static boolean requiresStandardInput(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return dev.erst.gridgrind.engine.runtime.SourceBackedPlanResolver.requiresStandardInput(
        request);
  }
}
