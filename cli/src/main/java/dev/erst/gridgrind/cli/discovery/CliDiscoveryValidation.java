package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;

/** Package-private validation helpers shared by CLI-owned discovery record constructors. */
final class CliDiscoveryValidation {
  private CliDiscoveryValidation() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static GridGrindProtocolVersion requireProtocolVersion(GridGrindProtocolVersion protocolVersion) {
    return Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
  }

  static WorkbookPlan requireRequestTemplate(WorkbookPlan requestTemplate) {
    return Objects.requireNonNull(requestTemplate, "requestTemplate must not be null");
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    return List.copyOf(values.stream().map(value -> requireNonBlank(value, fieldName)).toList());
  }

  static List<String> copyOptionalStrings(List<String> values) {
    if (values == null) {
      return List.of();
    }
    return List.copyOf(values.stream().map(value -> requireNonBlank(value, "value")).toList());
  }

  static List<TaskEntry> copyTaskEntries(List<TaskEntry> tasks, String fieldName) {
    return copyUnique(tasks, fieldName, TaskEntry::id);
  }

  static List<TaskPhase> copyTaskPhases(List<TaskPhase> phases, String fieldName) {
    return copyUnique(phases, fieldName, TaskPhase::label);
  }

  static List<TaskCapabilityRef> copyTaskCapabilityRefs(
      List<TaskCapabilityRef> capabilityRefs, String fieldName) {
    return copyUnique(capabilityRefs, fieldName, TaskCapabilityRef::qualifiedId);
  }

  static <E extends Enum<E>> List<E> copyEnumValues(
      List<E> values, String fieldName, java.util.function.Function<E, String> keyExtractor) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    return copyUnique(values, fieldName, keyExtractor);
  }

  static List<ShippedExampleEntry> copyExampleEntries(
      List<ShippedExampleEntry> entries, String fieldName) {
    return copyUnique(entries, fieldName, ShippedExampleEntry::id);
  }

  static List<TypeEntry> copyTypeEntries(List<TypeEntry> entries, String fieldName) {
    Objects.requireNonNull(entries, fieldName + " must not be null");
    return List.copyOf(
        entries.stream()
            .map(entry -> Objects.requireNonNull(entry, fieldName + " must not contain nulls"))
            .toList());
  }

  private static <T> List<T> copyUnique(
      List<T> values, String fieldName, Function<T, String> keyExtractor) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    Set<String> seen = new LinkedHashSet<>();
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      T nonNullValue = Objects.requireNonNull(value, fieldName + " must not contain nulls");
      String key = requireNonBlank(keyExtractor.apply(nonNullValue), fieldName);
      if (!seen.add(key)) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicate " + key);
      }
      copy.add(nonNullValue);
    }
    return List.copyOf(copy);
  }
}
