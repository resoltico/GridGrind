package dev.erst.gridgrind.cli.discovery;

/** One published field-metadata key carried by protocol-catalog lookup payloads. */
public record ProtocolCatalogFieldMetadataKey(String name, String meaning) {
  public ProtocolCatalogFieldMetadataKey {
    name = CliDiscoveryValidation.requireNonBlank(name, "name");
    meaning = CliDiscoveryValidation.requireNonBlank(meaning, "meaning");
  }
}
