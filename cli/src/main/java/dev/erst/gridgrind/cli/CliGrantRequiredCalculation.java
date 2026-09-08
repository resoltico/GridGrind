package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;

/** CLI grant requirement for strict post-mutation calculation before final publication. */
public record CliGrantRequiredCalculation() implements CliGrantAcceptancePolicy.Requirement {
  @Override
  public GridGrindHostAcceptancePolicy.Requirement toEngineRequirement() {
    return new GridGrindHostAcceptancePolicy.Requirement.RequireCalculation();
  }
}
