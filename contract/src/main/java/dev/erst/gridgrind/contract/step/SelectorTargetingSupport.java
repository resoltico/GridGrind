package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.ArrayList;
import java.util.List;

/** Shared selector-targeting helpers for workbook step validation registries. */
final class SelectorTargetingSupport {
  private SelectorTargetingSupport() {}

  static String humanTargetTypes(Class<? extends Selector>[] allowedTypes) {
    List<String> typeIds = new ArrayList<>();
    for (Class<? extends Selector> allowedType : allowedTypes) {
      typeIds.addAll(SelectorJsonSupport.typeIdsFor(allowedType));
    }
    if (typeIds.size() == 1) {
      return typeIds.getFirst();
    }
    StringBuilder builder = new StringBuilder();
    for (int index = 0; index < typeIds.size(); index++) {
      if (index > 0) {
        builder.append(index == typeIds.size() - 1 ? " or " : ", ");
      }
      builder.append(typeIds.get(index));
    }
    return builder.toString();
  }
}
