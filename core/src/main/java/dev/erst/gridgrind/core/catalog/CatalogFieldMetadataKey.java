package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One published field-metadata key carried by protocol-catalog lookup payloads. */
public record CatalogFieldMetadataKey(String name, String meaning) {
  public CatalogFieldMetadataKey {
    name = TextValidation.requireNonBlank(name, "name");
    meaning = TextValidation.requireNonBlank(meaning, "meaning");
  }
}
