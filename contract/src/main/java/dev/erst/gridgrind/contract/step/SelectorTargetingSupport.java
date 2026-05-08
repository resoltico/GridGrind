package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;

/** Shared selector-targeting helpers for workbook step validation registries. */
final class SelectorTargetingSupport {
  private SelectorTargetingSupport() {}

  @SafeVarargs
  static void requireTargetType(
      Selector target, String stepType, Class<? extends Selector>... allowedTypes) {
    for (Class<? extends Selector> allowedType : allowedTypes) {
      if (allowedType.isInstance(target)) {
        return;
      }
    }
    throw new IllegalArgumentException(
        stepType
            + " requires target type "
            + humanTargetTypes(allowedTypes)
            + " but got "
            + SelectorJsonSupport.displayName(target.getClass()));
  }

  static String humanTargetTypes(Class<? extends Selector>[] allowedTypes) {
    if (allowedTypes.length == 1) {
      return SelectorJsonSupport.displayName(allowedTypes[0]);
    }
    StringBuilder builder = new StringBuilder();
    for (int index = 0; index < allowedTypes.length; index++) {
      if (index > 0) {
        builder.append(index == allowedTypes.length - 1 ? " or " : ", ");
      }
      builder.append(SelectorJsonSupport.displayName(allowedTypes[index]));
    }
    return builder.toString();
  }
}
