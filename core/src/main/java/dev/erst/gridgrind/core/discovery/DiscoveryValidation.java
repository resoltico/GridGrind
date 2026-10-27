package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.catalog.TypeEntry;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;

/** Package-private validation helpers shared by Product-owned discovery record constructors. */
public final class DiscoveryValidation {
  private DiscoveryValidation() {}

  /** Requires the version carried by a discovery or command result. */
  public static ProtocolVersion requireProtocolVersion(ProtocolVersion protocolVersion) {
    return Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    return List.copyOf(
        values.stream().map(value -> TextValidation.requireNonBlank(value, fieldName)).toList());
  }

  static List<String> copyStringsAllowEmpty(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    return List.copyOf(
        values.stream().map(value -> TextValidation.requireNonBlank(value, fieldName)).toList());
  }

  /** Normalizes an absent list to empty and copies each present nonblank string. */
  public static List<String> copyOptionalStringsAllowEmpty(List<String> values, String fieldName) {
    if (values == null) {
      return List.of();
    }
    return copyStringsAllowEmpty(values, fieldName);
  }

  static java.util.Optional<String> normalizeOptionalString(
      java.util.Optional<String> value, String fieldName) {
    java.util.Optional<String> normalized =
        java.util.Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isPresent()) {
      normalized =
          java.util.Optional.of(
              TextValidation.requireNonBlank(normalized.orElseThrow(), fieldName));
    }
    return normalized;
  }

  static java.util.Optional<String> copyOptionalArbitraryString(
      java.util.Optional<String> value, String fieldName) {
    java.util.Optional<String> normalized =
        java.util.Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isEmpty()) {
      return java.util.Optional.empty();
    }
    return java.util.Optional.of(normalized.orElseThrow());
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

  static <T> List<T> copyUnique(
      List<T> values, String fieldName, Function<T, String> keyExtractor) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    Set<String> seen = new LinkedHashSet<>();
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      T nonNullValue = Objects.requireNonNull(value, fieldName + " must not contain nulls");
      String key = TextValidation.requireNonBlank(keyExtractor.apply(nonNullValue), fieldName);
      if (!seen.add(key)) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicate " + key);
      }
      copy.add(nonNullValue);
    }
    return List.copyOf(copy);
  }
}
