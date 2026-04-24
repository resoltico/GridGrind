package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

/** Shared selector wire-type metadata for root-level selector serialization. */
public final class SelectorJsonSupport {
  private static final List<Class<?>> SELECTOR_ROOTS =
      List.of(
          WorkbookSelector.class,
          SheetSelector.class,
          CellSelector.class,
          RangeSelector.class,
          RowBandSelector.class,
          ColumnBandSelector.class,
          TableSelector.class,
          TableRowSelector.class,
          TableCellSelector.class,
          NamedRangeSelector.class,
          DrawingObjectSelector.class,
          ChartSelector.class,
          PivotTableSelector.class);
  private static final Map<Class<?>, String> TYPE_IDS = buildTypeIds(SELECTOR_ROOTS);
  private static final java.util.Set<String> KNOWN_TYPE_IDS = buildKnownTypeIds(TYPE_IDS);

  private SelectorJsonSupport() {}

  /** Returns the canonical wire `type` discriminator for one concrete selector subtype. */
  public static String typeIdFor(Class<?> selectorType) {
    String typeId = TYPE_IDS.get(selectorType);
    if (typeId == null) {
      throw new IllegalArgumentException("Unsupported selector runtime type: " + selectorType);
    }
    return typeId;
  }

  /** Returns whether the supplied wire `type` is a shipped selector discriminator. */
  public static boolean isKnownTypeId(String typeId) {
    return KNOWN_TYPE_IDS.contains(typeId);
  }

  /** Returns the allowed wire `type` ids represented by one selector root or concrete subtype. */
  public static List<String> typeIdsFor(Class<?> selectorType) {
    JsonSubTypes jsonSubTypes = selectorType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes != null) {
      java.util.Set<String> typeIds = new LinkedHashSet<>();
      for (JsonSubTypes.Type subtype : jsonSubTypes.value()) {
        typeIds.add(subtype.name());
      }
      return List.copyOf(typeIds);
    }
    return List.of(typeIdFor(selectorType));
  }

  /** Returns whether one selector root or concrete subtype accepts the supplied wire id. */
  public static boolean supportsTypeId(Class<?> selectorType, String typeId) {
    return typeIdsFor(selectorType).contains(typeId);
  }

  /** Returns grouped selector-family metadata for the supplied selector roots or concrete types. */
  public static List<FamilyInfo> familyInfosFor(Iterable<Class<? extends Selector>> selectorTypes) {
    return StreamSupport.stream(selectorTypes.spliterator(), false)
        .collect(
            Collectors.groupingBy(
                SelectorJsonSupport::familyName,
                java.util.LinkedHashMap::new,
                Collectors.flatMapping(
                    selectorType -> typeIdsFor(selectorType).stream(),
                    Collectors.toCollection(LinkedHashSet::new))))
        .entrySet()
        .stream()
        .map(entry -> new FamilyInfo(entry.getKey(), List.copyOf(entry.getValue())))
        .toList();
  }

  /** Returns one stable human-readable selector-family name for the supplied selector type. */
  public static String familyName(Class<?> selectorType) {
    Class<?> enclosingClass = selectorType.getEnclosingClass();
    return enclosingClass == null ? selectorType.getSimpleName() : enclosingClass.getSimpleName();
  }

  /** Returns a stable human-readable summary of the supplied selector families and wire ids. */
  public static String familySummary(Iterable<Class<? extends Selector>> selectorTypes) {
    return familyInfosFor(selectorTypes).stream()
        .map(
            familyInfo -> familyInfo.family() + "(" + String.join(", ", familyInfo.typeIds()) + ")")
        .collect(java.util.stream.Collectors.joining("; "));
  }

  private static java.util.Set<String> buildKnownTypeIds(Map<Class<?>, String> typeIds) {
    return java.util.Set.copyOf(typeIds.values());
  }

  static Map<Class<?>, String> buildTypeIds(List<Class<?>> selectorRoots) {
    Map<Class<?>, String> typeIds = new ConcurrentHashMap<>();
    Map<String, Class<?>> typeIdOwners = new ConcurrentHashMap<>();
    for (Class<?> root : selectorRoots) {
      JsonSubTypes jsonSubTypes = root.getAnnotation(JsonSubTypes.class);
      if (jsonSubTypes == null) {
        throw new IllegalStateException("Selector root is missing @JsonSubTypes: " + root);
      }
      for (JsonSubTypes.Type subtype : jsonSubTypes.value()) {
        Class<?> previousOwner = typeIdOwners.putIfAbsent(subtype.name(), subtype.value());
        if (previousOwner != null && !previousOwner.equals(subtype.value())) {
          throw new IllegalStateException(
              "Duplicate selector type id '%s' for %s and %s"
                  .formatted(subtype.name(), previousOwner.getName(), subtype.value().getName()));
        }
        typeIds.put(subtype.value(), subtype.name());
      }
    }
    return Map.copyOf(typeIds);
  }

  /** One stable selector-family description with its allowed wire type ids. */
  public record FamilyInfo(String family, List<String> typeIds) {
    public FamilyInfo {
      Objects.requireNonNull(family, "family must not be null");
      if (family.isBlank()) {
        throw new IllegalArgumentException("family must not be blank");
      }
      Objects.requireNonNull(typeIds, "typeIds must not be null");
      if (typeIds.isEmpty()) {
        throw new IllegalArgumentException("typeIds must not be empty");
      }
      typeIds =
          typeIds.stream()
              .map(
                  typeId -> {
                    Objects.requireNonNull(typeId, "typeIds must not contain nulls");
                    if (typeId.isBlank()) {
                      throw new IllegalArgumentException("typeIds must not contain blank values");
                    }
                    return typeId;
                  })
              .toList();
    }
  }
}
