package dev.erst.gridgrind.cli.discovery;

/** One stable `--lookup` namespace form published by the protocol-catalog index. */
public record ProtocolCatalogLookupNamespace(String shape, String usage) {
  public ProtocolCatalogLookupNamespace {
    shape = CliDiscoveryValidation.requireNonBlank(shape, "shape");
    usage = CliDiscoveryValidation.requireNonBlank(usage, "usage");
  }
}
