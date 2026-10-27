package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;

/** JSON-serializable Product-owned catalog for built-in generated example requests. */
public record ShippedExampleCatalog(
    ProtocolVersion protocolVersion, List<ShippedExampleEntry> examples) {
  public ShippedExampleCatalog {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    examples = DiscoveryValidation.copyExampleEntries(examples, "examples");
  }
}
