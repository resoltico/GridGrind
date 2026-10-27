package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One stable shared note published once and referenced from catalog entries by id. */
public record CatalogNote(String id, String text) {
  public CatalogNote {
    id = TextValidation.requireNonBlank(id, "id");
    text = TextValidation.requireNonBlank(text, "text");
  }
}
