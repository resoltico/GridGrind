package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import java.util.List;

/** JSON-serializable CLI-owned catalog for built-in generated example requests. */
public record ShippedExampleCatalog(
    GridGrindProtocolVersion protocolVersion, List<ShippedExampleEntry> examples) {
  public ShippedExampleCatalog {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    examples = CliDiscoveryValidation.copyExampleEntries(examples, "examples");
  }
}
