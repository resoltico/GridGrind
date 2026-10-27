package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;

/** One published field-metadata key carried by protocol-catalog lookup payloads. */
public record ProtocolCatalogFieldMetadataKey(String name, String meaning) {
  public ProtocolCatalogFieldMetadataKey {
    name = DiscoveryValidation.requireNonBlank(name, "name");
    meaning = DiscoveryValidation.requireNonBlank(meaning, "meaning");
  }
}
