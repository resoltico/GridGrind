package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One published field-metadata key carried by protocol-catalog lookup payloads. */
public record ProtocolCatalogFieldMetadataKey(String name, String meaning) {
  public ProtocolCatalogFieldMetadataKey {
    name = TextValidation.requireNonBlank(name, "name");
    meaning = TextValidation.requireNonBlank(meaning, "meaning");
  }
}
