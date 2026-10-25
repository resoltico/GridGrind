package dev.erst.gridgrind.contract.catalog;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.function.Function;

/** Builds immutable descriptor maps while preserving duplicate checks and declaration order. */
final class CatalogDescriptorMaps {
  private CatalogDescriptorMaps() {}

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static <T, K, V> Map<K, V> uniqueMap(
      Iterable<T> values,
      Function<T, K> keyMapper,
      Function<T, V> valueMapper,
      String duplicatePrefix) {
    // Descriptor maps are assembled once during single-threaded catalog bootstrap, and
    // insertion order is part of the published catalog surface.
    Map<K, V> mapped = new LinkedHashMap<>();
    for (T value : values) {
      K key = keyMapper.apply(value);
      V duplicate = mapped.putIfAbsent(key, valueMapper.apply(value));
      if (duplicate != null) {
        throw new IllegalStateException(duplicatePrefix + key);
      }
    }
    return Map.copyOf(mapped);
  }
}
