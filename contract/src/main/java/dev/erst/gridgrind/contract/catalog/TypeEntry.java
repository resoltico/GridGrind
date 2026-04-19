package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** JSON-serializable type entry describing one request or nested union variant. */
public record TypeEntry(String id, String summary, List<FieldEntry> fields) {
  public TypeEntry {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    fields = CatalogRecordValidation.copyFieldEntries(fields, "fields");
  }

  /** Returns the field entry with the given name, or null when this type has no such field. */
  public FieldEntry field(String name) {
    Objects.requireNonNull(name, "name must not be null");
    return fields.stream().filter(field -> field.name().equals(name)).findFirst().orElse(null);
  }
}
