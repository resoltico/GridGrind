package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One stable qualified lookup namespace form published by the protocol-catalog index. */
public record ProtocolCatalogLookupNamespace(String shape, String usage) {
  public ProtocolCatalogLookupNamespace {
    shape = TextValidation.requireNonBlank(shape, "shape");
    usage = TextValidation.requireNonBlank(usage, "usage");
  }
}
