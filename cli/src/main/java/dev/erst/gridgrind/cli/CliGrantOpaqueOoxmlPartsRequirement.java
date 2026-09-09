package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.util.List;

/** CLI grant requirement for byte-identical preservation of explicitly named OOXML parts. */
public record CliGrantOpaqueOoxmlPartsRequirement(List<String> partNames)
    implements CliGrantAcceptancePolicy.Requirement {
  public CliGrantOpaqueOoxmlPartsRequirement {
    partNames = CliGrantAcceptancePolicy.copyValues(partNames, "partNames");
  }

  @Override
  public GridGrindHostAcceptancePolicy.Requirement toEngineRequirement() {
    return new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(partNames);
  }
}
