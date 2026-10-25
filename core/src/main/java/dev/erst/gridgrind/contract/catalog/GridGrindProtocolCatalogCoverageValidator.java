package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.catalog.gather.CatalogDuplicateFailures;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;

/** Coverage and ordering validation for the published protocol catalog. */
final class GridGrindProtocolCatalogCoverageValidator {
  private GridGrindProtocolCatalogCoverageValidator() {}

  static void validateFieldShapeGroupMappings(
      List<CatalogNestedTypeDescriptor> nestedTypeGroups,
      List<CatalogPlainTypeDescriptor> plainTypeDescriptors) {
    Set<Class<?>> descriptorNestedTypes =
        nestedTypeGroups.stream()
            .map(CatalogNestedTypeDescriptor::sealedType)
            .collect(java.util.stream.Collectors.toSet());
    Set<Class<?>> descriptorPlainTypes =
        plainTypeDescriptors.stream()
            .map(CatalogPlainTypeDescriptor::recordType)
            .collect(java.util.stream.Collectors.toSet());

    for (CatalogNestedTypeDescriptor descriptor : nestedTypeGroups) {
      CatalogFieldShapeRegistry.validateNestedTypeGroupMapping(
          descriptor.sealedType(), descriptor.group());
    }
    for (CatalogPlainTypeDescriptor descriptor : plainTypeDescriptors) {
      CatalogFieldShapeRegistry.validatePlainTypeGroupMapping(
          descriptor.recordType(), descriptor.group());
    }

    validateReverseGroupMappings(descriptorNestedTypes, descriptorPlainTypes);
  }

  static void validateReverseGroupMappings(
      Set<Class<?>> descriptorNestedTypes, Set<Class<?>> descriptorPlainTypes) {
    for (Class<?> registeredType : CatalogFieldShapeRegistry.registeredNestedTypes()) {
      if (!descriptorNestedTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape nested group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
    for (Class<?> registeredType : CatalogFieldShapeRegistry.registeredPlainTypes()) {
      if (!descriptorPlainTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape plain group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
  }

  static void validateCoverage(Class<?> sealedType, List<CatalogTypeDescriptor> descriptors) {
    validateCoverage(
        sealedType,
        toOrderedMap(
            descriptors,
            CatalogTypeDescriptor::recordType,
            descriptor -> descriptor.typeEntry().id(),
            "catalog descriptor"));
  }

  static void validateCoverage(Class<?> sealedType, Map<Class<?>, String> catalogIds) {
    if (sealedType.equals(WorkbookStep.class)) {
      validateWorkbookStepCoverage(sealedType, catalogIds);
      return;
    }
    CatalogTypeEntryFactory.discriminatorFieldFor(sealedType);
    Map<Class<?>, String> annotationIds = annotationIds(sealedType);
    validateCatalogRecords(catalogIds);
    validateCoveredTypes(sealedType, annotationIds, catalogIds);
    validateCoveredIds(annotationIds, catalogIds);
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static <T, K, V> Map<K, V> toOrderedMap(
      List<T> items, Function<T, K> keyFn, Function<T, V> valueFn, String label) {
    Map<K, V> result = new LinkedHashMap<>();
    for (T item : items) {
      K key = keyFn.apply(item);
      V value = valueFn.apply(item);
      if (result.containsKey(key)) {
        throw CatalogDuplicateFailures.duplicateEntryFailure(label, result.get(key), value);
      }
      result.put(key, value);
    }
    return result;
  }

  private static void validateWorkbookStepCoverage(
      Class<?> sealedType, Map<Class<?>, String> catalogIds) {
    Set<Class<?>> permitted =
        Arrays.stream(sealedType.getPermittedSubclasses())
            .collect(java.util.stream.Collectors.toCollection(java.util.LinkedHashSet::new));
    if (!permitted.equals(catalogIds.keySet())) {
      throw new IllegalStateException(
          "Catalog coverage mismatch for "
              + sealedType.getName()
              + ": permitted="
              + permitted
              + ", catalog="
              + catalogIds.keySet());
    }
  }

  private static Map<Class<?>, String> annotationIds(Class<?> sealedType) {
    return toOrderedMap(
        List.copyOf(ProtocolTypeMetadataSupport.typeIdsByClass(sealedType).entrySet()),
        Map.Entry::getKey,
        Map.Entry::getValue,
        "annotation subtype");
  }

  private static void validateCatalogRecords(Map<Class<?>, String> catalogIds) {
    for (Class<?> recordType : catalogIds.keySet()) {
      if (!recordType.isRecord()) {
        throw new IllegalStateException(
            "Catalog entry %s does not target a record type".formatted(recordType));
      }
    }
  }

  private static void validateCoveredTypes(
      Class<?> sealedType, Map<Class<?>, String> annotationIds, Map<Class<?>, String> catalogIds) {
    if (!annotationIds.keySet().equals(catalogIds.keySet())) {
      throw new IllegalStateException(
          "Catalog coverage mismatch for "
              + sealedType.getName()
              + ": annotated="
              + annotationIds.keySet()
              + ", catalog="
              + catalogIds.keySet());
    }
  }

  private static void validateCoveredIds(
      Map<Class<?>, String> annotationIds, Map<Class<?>, String> catalogIds) {
    for (Map.Entry<Class<?>, String> annotationEntry : annotationIds.entrySet()) {
      String catalogId = catalogIds.get(annotationEntry.getKey());
      if (!annotationEntry.getValue().equals(catalogId)) {
        throw new IllegalStateException(
            "Catalog id mismatch for "
                + annotationEntry.getKey().getName()
                + ": annotation="
                + annotationEntry.getValue()
                + ", catalog="
                + catalogId);
      }
    }
  }
}
