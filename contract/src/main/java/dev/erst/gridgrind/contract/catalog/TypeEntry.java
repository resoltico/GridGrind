package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** JSON-serializable type entry describing one request or nested union variant. */
public record TypeEntry(
    String id,
    String summary,
    List<FieldEntry> fields,
    List<TargetSelectorEntry> targetSelectors,
    String targetSelectorRule) {
  /**
   * Creates a type entry without target-selector metadata.
   *
   * <p>Use this overload for nested value types that are not step-addressable.
   */
  public TypeEntry(String id, String summary, List<FieldEntry> fields) {
    this(id, summary, fields, List.of(), null);
  }

  public TypeEntry {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    fields = CatalogRecordValidation.copyFieldEntries(fields, "fields");
    targetSelectors =
        CatalogRecordValidation.copyTargetSelectorEntries(targetSelectors, "targetSelectors");
    if (targetSelectorRule != null && targetSelectorRule.isBlank()) {
      throw new IllegalArgumentException("targetSelectorRule must not be blank");
    }
  }

  /** Returns the field entry with the given name, or empty when this type has no such field. */
  public Optional<FieldEntry> field(String name) {
    Objects.requireNonNull(name, "name must not be null");
    return fields.stream().filter(field -> field.name().equals(name)).findFirst();
  }
}
