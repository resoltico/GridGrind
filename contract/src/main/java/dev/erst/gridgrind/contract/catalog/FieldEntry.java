package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** JSON-serializable field descriptor for one record component. */
public record FieldEntry(
    String name, FieldRequirement requirement, FieldShape shape, List<String> enumValues) {
  public FieldEntry {
    name = CatalogRecordValidation.requireNonBlank(name, "name");
    Objects.requireNonNull(requirement, "requirement must not be null");
    Objects.requireNonNull(shape, "shape must not be null");
    enumValues = CatalogRecordValidation.copyStrings(enumValues, "enumValues");
  }
}
