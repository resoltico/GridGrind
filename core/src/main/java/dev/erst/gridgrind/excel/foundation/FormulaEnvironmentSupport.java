package dev.erst.gridgrind.excel.foundation;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;

/** Shared normalization helpers for formula-environment names, toolpacks, and UDF catalogs. */
public final class FormulaEnvironmentSupport {
  private FormulaEnvironmentSupport() {}

  /**
   * Copies one optional name-keyed list, treating {@code null} as an empty list while rejecting
   * null elements and case-insensitive duplicate names.
   */
  public static <T> List<T> copyOptionalDistinctNamedValues(
      List<T> values,
      String fieldName,
      String duplicateMessagePrefix,
      Function<T, String> nameExtractor) {
    if (values == null) {
      return List.of();
    }
    return copyDistinctNamedValues(values, fieldName, duplicateMessagePrefix, nameExtractor, false);
  }

  /**
   * Copies one required name-keyed list while rejecting null elements, empty lists, and
   * case-insensitive duplicate names.
   */
  public static <T> List<T> copyRequiredDistinctNamedValues(
      List<T> values,
      String fieldName,
      String emptyMessage,
      String duplicateMessagePrefix,
      Function<T, String> nameExtractor) {
    List<T> copy =
        copyDistinctNamedValues(values, fieldName, duplicateMessagePrefix, nameExtractor, true);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(emptyMessage);
    }
    return copy;
  }

  /**
   * Rejects case-insensitive duplicate names across every member of every already-normalized group.
   */
  public static <G, M> void requireDistinctNestedNames(
      List<G> groups,
      Function<G, List<M>> membersExtractor,
      Function<M, String> nameExtractor,
      String duplicateMessagePrefix) {
    Objects.requireNonNull(groups, "groups must not be null");
    Objects.requireNonNull(membersExtractor, "membersExtractor must not be null");
    Objects.requireNonNull(nameExtractor, "nameExtractor must not be null");
    Objects.requireNonNull(duplicateMessagePrefix, "duplicateMessagePrefix must not be null");
    Set<String> seen = new LinkedHashSet<>();
    for (G group : groups) {
      for (M member : membersExtractor.apply(group)) {
        String name = nameExtractor.apply(member);
        if (!seen.add(normalizeName(name))) {
          throw new IllegalArgumentException(duplicateMessagePrefix + name);
        }
      }
    }
  }

  private static <T> List<T> copyDistinctNamedValues(
      List<T> values,
      String fieldName,
      String duplicateMessagePrefix,
      Function<T, String> nameExtractor,
      boolean requireNonNullList) {
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    Objects.requireNonNull(duplicateMessagePrefix, "duplicateMessagePrefix must not be null");
    Objects.requireNonNull(nameExtractor, "nameExtractor must not be null");
    if (requireNonNullList) {
      Objects.requireNonNull(values, fieldName + " must not be null");
    }
    List<T> copy = copyValues(values, fieldName);
    Set<String> seen = new LinkedHashSet<>();
    for (T value : copy) {
      String name = nameExtractor.apply(value);
      if (!seen.add(normalizeName(name))) {
        throw new IllegalArgumentException(duplicateMessagePrefix + name);
      }
    }
    return copy;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  private static String normalizeName(String value) {
    return Objects.requireNonNull(value, "name must not be null").toUpperCase(Locale.ROOT);
  }
}
