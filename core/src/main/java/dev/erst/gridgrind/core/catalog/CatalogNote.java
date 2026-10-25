package dev.erst.gridgrind.core.catalog;

/** One stable shared note published once and referenced from catalog entries by id. */
public record CatalogNote(String id, String text) {
  public CatalogNote {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    text = CatalogRecordValidation.requireNonBlank(text, "text");
  }
}
