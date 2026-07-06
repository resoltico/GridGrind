package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** JSON-serializable type entry describing one request or nested union variant. */
public record TypeEntry(
    String id,
    String summary,
    List<FieldEntry> fields,
    List<TargetSelectorEntry> targetSelectors,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> targetSelectorRule,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<String> noteRefs,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ProtocolStepTemplate> stepTemplate) {
  /**
   * Creates a type entry without target-selector metadata.
   *
   * <p>Use this overload for nested value types that are not step-addressable.
   */
  public TypeEntry(String id, String summary, List<FieldEntry> fields) {
    this(id, summary, fields, List.of(), Optional.empty(), List.of(), Optional.empty());
  }

  /** Creates a type entry with target-selector metadata but without a step template. */
  public TypeEntry(
      String id,
      String summary,
      List<FieldEntry> fields,
      List<TargetSelectorEntry> targetSelectors,
      Optional<String> targetSelectorRule) {
    this(id, summary, fields, targetSelectors, targetSelectorRule, List.of(), Optional.empty());
  }

  /** Creates a type entry with target selectors and shared-note references. */
  public TypeEntry(
      String id,
      String summary,
      List<FieldEntry> fields,
      List<TargetSelectorEntry> targetSelectors,
      Optional<String> targetSelectorRule,
      List<String> noteRefs) {
    this(id, summary, fields, targetSelectors, targetSelectorRule, noteRefs, Optional.empty());
  }

  public TypeEntry {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    fields = CatalogRecordValidation.copyFieldEntries(fields, "fields");
    targetSelectors =
        CatalogRecordValidation.copyTargetSelectorEntries(targetSelectors, "targetSelectors");
    Objects.requireNonNull(targetSelectorRule, "targetSelectorRule must not be null");
    if (targetSelectorRule.isPresent() && targetSelectorRule.orElseThrow().isBlank()) {
      throw new IllegalArgumentException("targetSelectorRule must not be blank");
    }
    noteRefs = Objects.requireNonNullElseGet(noteRefs, List::of);
    noteRefs = CatalogRecordValidation.copyUniqueStrings(noteRefs, "noteRefs");
    Objects.requireNonNull(stepTemplate, "stepTemplate must not be null");
  }

  /** Returns the field entry with the given name, or empty when this type has no such field. */
  public Optional<FieldEntry> field(String name) {
    Objects.requireNonNull(name, "name must not be null");
    return fields.stream().filter(field -> field.name().equals(name)).findFirst();
  }
}
