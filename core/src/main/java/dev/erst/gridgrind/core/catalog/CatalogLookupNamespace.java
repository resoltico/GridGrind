package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One stable qualified lookup namespace form published by the protocol-catalog index. */
public record CatalogLookupNamespace(String shape, String usage) {
  public CatalogLookupNamespace {
    shape = TextValidation.requireNonBlank(shape, "shape");
    usage = TextValidation.requireNonBlank(usage, "usage");
  }
}
