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

  static List<ShippedExampleEntry> copyExampleEntries(
      List<ShippedExampleEntry> entries, String fieldName) {
    Objects.requireNonNull(entries, fieldName + " must not be null");
    List<ShippedExampleEntry> copy = new java.util.ArrayList<>(entries.size());
    for (ShippedExampleEntry entry : entries) {
      copy.add(Objects.requireNonNull(entry, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(ShippedExampleEntry::id, fieldName))
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

  static List<TaskEntry> copyTaskEntries(List<TaskEntry> tasks, String fieldName) {
    Objects.requireNonNull(tasks, fieldName + " must not be null");
    List<TaskEntry> copy = new java.util.ArrayList<>(tasks.size());
    for (TaskEntry task : tasks) {
      copy.add(Objects.requireNonNull(task, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(TaskEntry::id, fieldName))
        .toList();
  }

  static List<TaskPhase> copyTaskPhases(List<TaskPhase> phases, String fieldName) {
    Objects.requireNonNull(phases, fieldName + " must not be null");
    List<TaskPhase> copy = new java.util.ArrayList<>(phases.size());
    for (TaskPhase phase : phases) {
      copy.add(Objects.requireNonNull(phase, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(TaskPhase::label, fieldName))
        .toList();
  }

  static List<TaskCapabilityRef> copyTaskCapabilityRefs(
      List<TaskCapabilityRef> capabilityRefs, String fieldName) {
    Objects.requireNonNull(capabilityRefs, fieldName + " must not be null");
    List<TaskCapabilityRef> copy = new java.util.ArrayList<>(capabilityRefs.size());
    for (TaskCapabilityRef capabilityRef : capabilityRefs) {
      copy.add(Objects.requireNonNull(capabilityRef, fieldName + " must not contain nulls"));
    }
    return copy.stream()
        .gather(CatalogGatherers.toOrderedUniqueOrThrow(TaskCapabilityRef::qualifiedId, fieldName))
        .toList();
  }
}
