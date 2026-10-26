package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.plan.Selector;
import java.util.List;
import java.util.Optional;
import java.util.function.Function;

/** Resolves field references against groups derived from model metadata. */
final class CatalogGroups {
  private CatalogGroups() {}

  static Optional<FieldShape> fieldShape(Class<?> type) {
    if (type == Selector.class) {
      return Optional.of(
          new FieldShape.NestedTypeGroupUnionRef(
              CatalogTypes.NESTED_TYPE_GROUPS.stream()
                  .filter(group -> Selector.class.isAssignableFrom(group.sealedType()))
                  .map(CatalogNestedTypeDescriptor::group)
                  .toList()));
    }
    Optional<String> topLevel =
        groupFor(
            CatalogTypes.TOP_LEVEL_GROUPS,
            type,
            CatalogTopLevelTypeDescriptorGroup::sealedType,
            CatalogTopLevelTypeDescriptorGroup::group);
    if (topLevel.isPresent()) {
      return topLevel.map(FieldShape.TopLevelTypeSetRef::new);
    }
    Optional<String> nested =
        groupFor(
            CatalogTypes.NESTED_TYPE_GROUPS,
            type,
            CatalogNestedTypeDescriptor::sealedType,
            CatalogNestedTypeDescriptor::group);
    if (nested.isPresent()) {
      return nested.map(FieldShape.NestedTypeGroupRef::new);
    }
    return type.isAnnotationPresent(CatalogType.class)
            && type != dev.erst.gridgrind.core.plan.WorkbookPlan.class
        ? Optional.of(new FieldShape.PlainTypeGroupRef(CatalogTypes.plainGroupId(type)))
        : Optional.empty();
  }

  private static <T> Optional<String> groupFor(
      List<T> groups, Class<?> type, Function<T, Class<?>> root, Function<T, String> name) {
    T selected = null;
    for (T candidate : groups) {
      Class<?> candidateType = root.apply(candidate);
      if (!candidateType.isAssignableFrom(type)) {
        continue;
      }
      if (candidateType == type) {
        return Optional.of(name.apply(candidate));
      }
      if (selected == null || root.apply(selected).isAssignableFrom(candidateType)) {
        selected = candidate;
      } else if (!candidateType.isAssignableFrom(root.apply(selected))) {
        throw new IllegalStateException(
            "Ambiguous catalog field groups for "
                + type.getName()
                + ": "
                + root.apply(selected).getName()
                + " and "
                + candidateType.getName());
      }
    }
    return Optional.ofNullable(selected).map(name);
  }
}
