package dev.erst.gridgrind.protocol.catalog;

import dev.erst.gridgrind.protocol.catalog.gather.CatalogGatherers;
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
    List<TypeEntry> copy = List.copyOf(entries);
    for (TypeEntry entry : copy) {
      Objects.requireNonNull(entry, fieldName + " must not contain nulls");
    }
    return copy;
  }

  static List<NestedTypeGroup> copyGroups(List<NestedTypeGroup> groups, String fieldName) {
    Objects.requireNonNull(groups, fieldName + " must not be null");
    List<NestedTypeGroup> copy = List.copyOf(groups);
    for (NestedTypeGroup group : copy) {
      Objects.requireNonNull(group, fieldName + " must not contain nulls");
    }
    return copy;
  }

  static List<PlainTypeGroup> copyPlainGroups(List<PlainTypeGroup> groups, String fieldName) {
    Objects.requireNonNull(groups, fieldName + " must not be null");
    List<PlainTypeGroup> copy = List.copyOf(groups);
    for (PlainTypeGroup group : copy) {
      Objects.requireNonNull(group, fieldName + " must not contain nulls");
    }
    return copy;
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      requireNonBlank(value, fieldName);
    }
    return copy;
  }

  static List<FieldEntry> copyFieldEntries(List<FieldEntry> fields, String fieldName) {
    Objects.requireNonNull(fields, fieldName + " must not be null");
    List<FieldEntry> copy = List.copyOf(fields);
    copy.forEach(field -> Objects.requireNonNull(field, fieldName + " must not contain nulls"));
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(FieldEntry::name, fieldName))
        .toList();
  }
}
