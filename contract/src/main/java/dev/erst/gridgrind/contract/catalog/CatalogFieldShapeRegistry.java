package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.selector.Selector;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

/** Registry-backed field-shape group resolution for protocol-catalog metadata. */
final class CatalogFieldShapeRegistry {
  private CatalogFieldShapeRegistry() {}

  static Optional<FieldShape> groupedFieldShape(Class<?> classType) {
    return resolvedGroupList(nestedFieldShapeUnions(), classType)
        .<FieldShape>map(FieldShape.NestedTypeGroupUnionRef::new)
        .or(
            () ->
                resolvedGroup(topLevelFieldShapeTypeSets(), classType)
                    .map(FieldShape.TopLevelTypeSetRef::new))
        .or(
            () ->
                resolvedGroup(nestedFieldShapeGroups(), classType)
                    .map(FieldShape.NestedTypeGroupRef::new))
        .or(
            () ->
                resolvedGroup(plainFieldShapeGroups(), classType)
                    .map(FieldShape.PlainTypeGroupRef::new));
  }

  static Optional<String> lookupAssignableGroup(Map<Class<?>, String> groups, Class<?> classType) {
    return groups.entrySet().stream()
        .filter(
            entry ->
                entry.getKey().isAssignableFrom(classType) && !entry.getKey().equals(classType))
        .reduce(
            (resolved, candidate) ->
                selectMoreSpecificGroup(
                    classType, resolved, candidate, "Ambiguous catalog field group mapping for "))
        .map(Map.Entry::getValue)
        .map(Optional::of)
        .orElseGet(Optional::empty);
  }

  static Optional<List<String>> lookupAssignableGroupList(
      Map<Class<?>, List<String>> groups, Class<?> classType) {
    return groups.entrySet().stream()
        .filter(
            entry ->
                entry.getKey().isAssignableFrom(classType) && !entry.getKey().equals(classType))
        .reduce(
            (resolved, candidate) ->
                selectMoreSpecificGroup(
                    classType,
                    resolved,
                    candidate,
                    "Ambiguous catalog field group-list mapping for "))
        .map(Map.Entry::getValue)
        .map(Optional::of)
        .orElseGet(Optional::empty);
  }

  static Set<Class<?>> registeredNestedTypes() {
    return nestedFieldShapeGroups().keySet();
  }

  static Set<Class<?>> registeredPlainTypes() {
    return plainFieldShapeGroups().keySet();
  }

  static void validateNestedTypeGroupMapping(Class<?> sealedType, String expectedGroup) {
    String mappedGroup = nestedFieldShapeGroups().get(sealedType);
    if (!expectedGroup.equals(mappedGroup)) {
      throw new IllegalStateException(
          "Field-shape nested group mapping mismatch for "
              + sealedType.getName()
              + ": expected="
              + expectedGroup
              + ", mapped="
              + mappedGroup);
    }
  }

  static void validatePlainTypeGroupMapping(Class<?> recordType, String expectedGroup) {
    String mappedGroup = plainFieldShapeGroups().get(recordType);
    if (!expectedGroup.equals(mappedGroup)) {
      throw new IllegalStateException(
          "Field-shape plain group mapping mismatch for "
              + recordType.getName()
              + ": expected="
              + expectedGroup
              + ", mapped="
              + mappedGroup);
    }
  }

  private static Optional<String> resolvedGroup(Map<Class<?>, String> groups, Class<?> classType) {
    String exact = groups.get(classType);
    return exact != null ? Optional.of(exact) : lookupAssignableGroup(groups, classType);
  }

  private static Optional<List<String>> resolvedGroupList(
      Map<Class<?>, List<String>> groups, Class<?> classType) {
    List<String> exact = groups.get(classType);
    return exact != null ? Optional.of(exact) : lookupAssignableGroupList(groups, classType);
  }

  private static <T> Map.Entry<Class<?>, T> selectMoreSpecificGroup(
      Class<?> classType,
      Map.Entry<Class<?>, T> resolved,
      Map.Entry<Class<?>, T> candidate,
      String errorPrefix) {
    Class<?> resolvedBase = resolved.getKey();
    Class<?> candidateBase = candidate.getKey();
    if (resolvedBase.isAssignableFrom(candidateBase)) {
      return candidate;
    }
    if (candidateBase.isAssignableFrom(resolvedBase)) {
      return resolved;
    }
    throw new IllegalStateException(
        errorPrefix
            + classType.getName()
            + ": "
            + resolvedBase.getName()
            + " and "
            + candidateBase.getName());
  }

  private static Map<Class<?>, String> nestedFieldShapeGroups() {
    return CatalogDescriptorMaps.uniqueMap(
        GridGrindProtocolCatalogNestedTypeGroups.NESTED_TYPE_GROUPS,
        CatalogNestedTypeDescriptor::sealedType,
        CatalogNestedTypeDescriptor::group,
        "Duplicate nested catalog field-shape group for ");
  }

  private static Map<Class<?>, List<String>> nestedFieldShapeUnions() {
    return Map.of(
        Selector.class,
        GridGrindProtocolCatalogNestedTypeGroups.NESTED_TYPE_GROUPS.stream()
            .filter(descriptor -> Selector.class.isAssignableFrom(descriptor.sealedType()))
            .map(CatalogNestedTypeDescriptor::group)
            .toList());
  }

  private static Map<Class<?>, String> topLevelFieldShapeTypeSets() {
    return CatalogDescriptorMaps.uniqueMap(
        GridGrindProtocolCatalogTypeDescriptors.TOP_LEVEL_GROUPS,
        CatalogTopLevelTypeDescriptorGroup::sealedType,
        CatalogTopLevelTypeDescriptorGroup::group,
        "Duplicate top-level catalog field-shape group for ");
  }

  private static Map<Class<?>, String> plainFieldShapeGroups() {
    return CatalogDescriptorMaps.uniqueMap(
        GridGrindProtocolCatalogPlainTypeDescriptors.PLAIN_TYPE_DESCRIPTORS,
        CatalogPlainTypeDescriptor::recordType,
        CatalogPlainTypeDescriptor::group,
        "Duplicate plain catalog field-shape group for ");
  }
}
