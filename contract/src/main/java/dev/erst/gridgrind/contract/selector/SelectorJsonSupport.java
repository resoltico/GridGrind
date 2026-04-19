package dev.erst.gridgrind.contract.selector;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/** Shared selector wire-type metadata for root-level selector serialization. */
final class SelectorJsonSupport {
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

  private SelectorJsonSupport() {}

  static String typeIdFor(Class<?> selectorType) {
    String typeId = TYPE_IDS.get(selectorType);
    if (typeId == null) {
      throw new IllegalArgumentException("Unsupported selector runtime type: " + selectorType);
    }
    return typeId;
  }

  static Map<Class<?>, String> buildTypeIds(List<Class<?>> selectorRoots) {
    Map<Class<?>, String> typeIds = new ConcurrentHashMap<>();
    for (Class<?> root : selectorRoots) {
      JsonSubTypes jsonSubTypes = root.getAnnotation(JsonSubTypes.class);
      if (jsonSubTypes == null) {
        throw new IllegalStateException("Selector root is missing @JsonSubTypes: " + root);
      }
      for (JsonSubTypes.Type subtype : jsonSubTypes.value()) {
        typeIds.put(subtype.value(), subtype.name());
      }
    }
    return Map.copyOf(typeIds);
  }
}
