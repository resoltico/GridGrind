package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;

/** Shared helpers for exact enum-to-enum lookup tables across workbook bridge seams. */
public final class ExcelEnumMappingSupport {
  private ExcelEnumMappingSupport() {}

  /** Builds an immutable enum-keyed map and fails fast if any enum constant lacks a mapping. */
  public static <K extends Enum<K>, V> Map<K, V> exactEnumMap(
      Class<K> keyType, String mappingName, Map<K, V> values) {
    Objects.requireNonNull(keyType, "keyType must not be null");
    Objects.requireNonNull(mappingName, "mappingName must not be null");
    Objects.requireNonNull(values, "values must not be null");
    K[] enumConstants = keyType.getEnumConstants();
    Set<K> expectedKeys = Set.copyOf(Arrays.asList(enumConstants));
    if (!values.keySet().equals(expectedKeys)) {
      throw new IllegalStateException(
          mappingName
              + " must cover every "
              + keyType.getSimpleName()
              + " constant; expected keys "
              + expectedKeys
              + " but found "
              + values.keySet());
    }
    return Map.copyOf(values);
  }

  /**
   * Reverses an exact enum mapping and fails fast when the reverse side is incomplete or ambiguous.
   */
  public static <K extends Enum<K>, V extends Enum<V>> Map<V, K> reverseExactEnumMap(
      Class<V> keyType, String mappingName, Map<K, V> forwardValues) {
    Objects.requireNonNull(keyType, "keyType must not be null");
    Objects.requireNonNull(mappingName, "mappingName must not be null");
    Objects.requireNonNull(forwardValues, "forwardValues must not be null");
    V[] enumConstants = keyType.getEnumConstants();
    Map<V, K> reversed =
        forwardValues.entrySet().stream()
            .collect(
                Collectors.toMap(
                    Map.Entry::getValue,
                    Map.Entry::getKey,
                    (left, right) -> {
                      throw new IllegalStateException(
                          mappingName + " maps " + left + " and " + right + " to the same target");
                    }));
    Set<V> expectedKeys = Set.copyOf(Arrays.asList(enumConstants));
    if (!reversed.keySet().equals(expectedKeys)) {
      throw new IllegalStateException(
          mappingName
              + " must cover every "
              + keyType.getSimpleName()
              + " constant; expected keys "
              + expectedKeys
              + " but found "
              + reversed.keySet());
    }
    return Map.copyOf(reversed);
  }

  /** Returns one mapped value or throws a caller-supplied unsupported-value error message. */
  public static <K, V> V requireMappedValue(Map<K, V> mapping, K key, String description) {
    Objects.requireNonNull(mapping, "mapping must not be null");
    V value = mapping.get(key);
    if (value == null) {
      throw new IllegalArgumentException("Unsupported " + description + ": " + key);
    }
    return value;
  }
}
