package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.catalog.gather.CatalogGatherers;
import java.util.List;
import java.util.Objects;

/** Package-private validation helpers shared by catalog record compact constructors. */
final class CatalogRecordValidation {
  private CatalogRecordValidation() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static List<TypeEntry> copyEntries(List<TypeEntry> entries, String fieldName) {
    Objects.requireNonNull(entries, fieldName + " must not be null");
    List<TypeEntry> copy = new java.util.ArrayList<>(entries.size());
    for (TypeEntry entry : entries) {
      copy.add(Objects.requireNonNull(entry, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  static List<CatalogNote> copyNotes(List<CatalogNote> notes, String fieldName) {
    Objects.requireNonNull(notes, fieldName + " must not be null");
    List<CatalogNote> copy = new java.util.ArrayList<>(notes.size());
    for (CatalogNote note : notes) {
      copy.add(Objects.requireNonNull(note, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(CatalogNote::id, fieldName))
        .toList();
  }

  static List<NestedTypeGroup> copyGroups(List<NestedTypeGroup> groups, String fieldName) {
    Objects.requireNonNull(groups, fieldName + " must not be null");
    List<NestedTypeGroup> copy = new java.util.ArrayList<>(groups.size());
    for (NestedTypeGroup group : groups) {
      copy.add(Objects.requireNonNull(group, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  static List<PlainTypeGroup> copyPlainGroups(List<PlainTypeGroup> groups, String fieldName) {
    Objects.requireNonNull(groups, fieldName + " must not be null");
    List<PlainTypeGroup> copy = new java.util.ArrayList<>(groups.size());
    for (PlainTypeGroup group : groups) {
      copy.add(Objects.requireNonNull(group, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = new java.util.ArrayList<>(values.size());
    for (String value : values) {
      copy.add(requireNonBlank(value, fieldName));
    }
    return List.copyOf(copy);
  }

  static List<String> copyUniqueStrings(List<String> values, String fieldName) {
    return copyStrings(values, fieldName).stream()
        .gather(
            CatalogGatherers.toOrderedUniqueOrThrow(
                java.util.function.Function.identity(), fieldName))
        .toList();
  }

  static List<EnumValueDocEntry> copyEnumValueDocs(List<EnumValueDocEntry> docs, String fieldName) {
    Objects.requireNonNull(docs, fieldName + " must not be null");
    List<EnumValueDocEntry> copy = new java.util.ArrayList<>(docs.size());
    for (EnumValueDocEntry doc : docs) {
      copy.add(Objects.requireNonNull(doc, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(EnumValueDocEntry::value, fieldName))
        .toList();
  }

  static List<FieldEntry> copyFieldEntries(List<FieldEntry> fields, String fieldName) {
    Objects.requireNonNull(fields, fieldName + " must not be null");
    List<FieldEntry> copy = new java.util.ArrayList<>(fields.size());
    for (FieldEntry field : fields) {
      copy.add(Objects.requireNonNull(field, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(FieldEntry::name, fieldName))
        .toList();
  }

  static List<CatalogProjectedField> copyProjectedFields(
      List<CatalogProjectedField> projectedFields, String fieldName) {
    Objects.requireNonNull(projectedFields, fieldName + " must not be null");
    List<CatalogProjectedField> copy = new java.util.ArrayList<>(projectedFields.size());
    for (CatalogProjectedField projectedField : projectedFields) {
      copy.add(Objects.requireNonNull(projectedField, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(CatalogProjectedField::name, fieldName))
        .toList();
  }

  static List<TargetSelectorEntry> copyTargetSelectorEntries(
      List<TargetSelectorEntry> targetSelectors, String fieldName) {
    Objects.requireNonNull(targetSelectors, fieldName + " must not be null");
    List<TargetSelectorEntry> copy = new java.util.ArrayList<>(targetSelectors.size());
    for (TargetSelectorEntry targetSelector : targetSelectors) {
      copy.add(Objects.requireNonNull(targetSelector, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(TargetSelectorEntry::family, fieldName))
        .toList();
  }
}
