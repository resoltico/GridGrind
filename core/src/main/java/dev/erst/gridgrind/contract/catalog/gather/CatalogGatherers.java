package dev.erst.gridgrind.contract.catalog.gather;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.contract.catalog.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.contract.catalog.CatalogIgnored;
import dev.erst.gridgrind.contract.catalog.FieldEntry;
import java.lang.reflect.RecordComponent;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Gatherer;

/** Small, domain-specific gatherers used by protocol-catalog construction. */
public final class CatalogGatherers {
  private CatalogGatherers() {}

  /** Emits the encounter-ordered items while rejecting duplicate catalog keys. */
  public static <T, K> Gatherer<T, ?, T> toOrderedUniqueOrThrow(
      Function<? super T, ? extends K> keyFunction, String label) {
    return toOrderedUniqueOrThrow(keyFunction, Function.identity(), label);
  }

  /** Emits encounter-ordered items while rejecting duplicates with projected failure values. */
  public static <T, K> Gatherer<T, ?, T> toOrderedUniqueOrThrow(
      Function<? super T, ? extends K> keyFunction,
      Function<? super T, ?> duplicateValueFunction,
      String label) {
    Objects.requireNonNull(keyFunction, "keyFunction must not be null");
    Objects.requireNonNull(duplicateValueFunction, "duplicateValueFunction must not be null");
    Objects.requireNonNull(label, "label must not be null");
    if (label.isBlank()) {
      throw new IllegalArgumentException("label must not be blank");
    }
    return Gatherer.ofSequential(
        LinkedHashMap<K, T>::new,
        (state, item, downstream) -> {
          K key = keyFunction.apply(item);
          T previous = state.putIfAbsent(key, item);
          if (previous != null) {
            throw CatalogDuplicateFailures.duplicateEntryFailure(
                label, duplicateValueFunction.apply(previous), duplicateValueFunction.apply(item));
          }
          downstream.push(item);
          return true;
        });
  }

  /** Expands reflected record components into enriched protocol-catalog field entries. */
  public static Gatherer<RecordComponent, ?, FieldEntry> expandFieldsWithMetadata(
      Set<String> optionalFields) {
    return expandFieldsWithMetadata(optionalFields, Map.of());
  }

  /** Expands reflected record components into enriched protocol-catalog field entries. */
  public static Gatherer<RecordComponent, ?, FieldEntry> expandFieldsWithMetadata(
      Set<String> optionalFields, Map<String, List<String>> projectedFieldsByName) {
    Objects.requireNonNull(optionalFields, "optionalFields must not be null");
    Objects.requireNonNull(projectedFieldsByName, "projectedFieldsByName must not be null");
    Set<String> optionalFieldSet = Set.copyOf(optionalFields);
    Map<String, List<String>> projectedFieldMap = Map.copyOf(projectedFieldsByName);
    return Gatherer.ofSequential(
        () -> new FieldMetadataState(optionalFieldSet, projectedFieldMap),
        (state, component, downstream) -> {
          if (component.isAnnotationPresent(CatalogIgnored.class)
              || component.getAccessor().isAnnotationPresent(CatalogIgnored.class)
              || component.getAccessor().isAnnotationPresent(JsonIgnore.class)) {
            return true;
          }
          downstream.push(
              CatalogFieldMetadataSupport.fieldEntry(
                  component, state.optionalFields(), state.projectedFieldsByName()));
          return true;
        });
  }

  private record FieldMetadataState(
      Set<String> optionalFields, Map<String, List<String>> projectedFieldsByName) {}
}
