package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;

/** One stable qualified lookup namespace form published by the protocol-catalog index. */
public record ProtocolCatalogLookupNamespace(String shape, String usage) {
  public ProtocolCatalogLookupNamespace {
    shape = DiscoveryValidation.requireNonBlank(shape, "shape");
    usage = DiscoveryValidation.requireNonBlank(usage, "usage");
  }
}
