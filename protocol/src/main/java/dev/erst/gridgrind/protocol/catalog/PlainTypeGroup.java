package dev.erst.gridgrind.protocol.catalog;

import java.util.Objects;

/** JSON-serializable named group describing one plain record type with its field shape. */
public record PlainTypeGroup(String group, TypeEntry type) {
  public PlainTypeGroup {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    Objects.requireNonNull(type, "type must not be null");
  }
}
