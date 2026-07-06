package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.gather.CatalogGatherers;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.contract.step.WorkbookStepTargeting;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/** Builds catalog descriptor payloads without forcing GridGrindProtocolCatalog initialization. */
final class CatalogTypeEntryFactory {
  private CatalogTypeEntryFactory() {}

  static CatalogNestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
    return new CatalogNestedTypeDescriptor(
        group, discriminatorFieldFor(sealedType), sealedType, typeDescriptors);
  }

  static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return new CatalogPlainTypeDescriptor(group, recordType, id, summary, optionalFields);
  }

  static CatalogPlainTypeDescriptor plainTypeDescriptorWithNotes(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields,
      List<String> noteRefs) {
    return new CatalogPlainTypeDescriptor(group, recordType, id, summary, optionalFields, noteRefs);
  }

  static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new CatalogTypeDescriptor(recordType, id, summary, List.of(optionalFields));
  }

  static CatalogTypeDescriptor descriptorWithNotes(
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> noteRefs,
      String... optionalFields) {
    return new CatalogTypeDescriptor(
        recordType, id, summary, List.of(optionalFields), noteRefs, List.of());
  }

  static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields,
      CatalogProjectedField... projectedFields) {
    return new CatalogTypeDescriptor(
        recordType, id, summary, optionalFields, List.of(), List.of(projectedFields));
  }

  static TypeEntry typeEntry(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    return typeEntry(recordType, id, summary, optionalFields, List.of(), List.of());
  }

  static TypeEntry typeEntry(
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields,
      List<String> noteRefs) {
    return typeEntry(recordType, id, summary, optionalFields, noteRefs, List.of());
  }

  static TypeEntry typeEntry(
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields,
      List<String> noteRefs,
      List<CatalogProjectedField> projectedFields) {
    Optional<WorkbookStepTargeting.TargetSurface> targetSurface =
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(recordType);
    return new TypeEntry(
        canonicalTypeId(recordType, id),
        summary,
        fieldEntries(recordType, optionalFields, projectedFields),
        TypeEntryTargetingSupport.targetSelectorEntries(targetSurface),
        targetSurface.flatMap(WorkbookStepTargeting.TargetSurface::rule),
        noteRefs);
  }

  static List<String> requiredFields(
      Class<? extends Record> recordType, List<String> optionalFields) {
    return requiredFields(recordType, optionalFields, List.of());
  }

  static List<String> requiredFields(
      Class<? extends Record> recordType,
      List<String> optionalFields,
      List<CatalogProjectedField> projectedFields) {
    List<String> recordFields = recordFields(recordType);
    for (String optionalField : optionalFields) {
      if (!recordFields.contains(optionalField)) {
        throw new IllegalStateException(
            "Catalog optional field '%s' does not exist on %s"
                .formatted(optionalField, recordType.getName()));
      }
    }
    for (CatalogProjectedField projectedField : projectedFields) {
      if (!recordFields.contains(projectedField.name())) {
        throw new IllegalStateException(
            "Catalog projected field '%s' does not exist on %s"
                .formatted(projectedField.name(), recordType.getName()));
      }
    }
    Set<String> optionalFieldSet =
        Stream.concat(
                optionalFields.stream(), projectedFields.stream().map(CatalogProjectedField::name))
            .collect(Collectors.toUnmodifiableSet());
    return recordFields.stream().filter(field -> !optionalFieldSet.contains(field)).toList();
  }

  static String discriminatorFieldFor(Class<?> sealedType) {
    JsonTypeInfo jsonTypeInfo = sealedType.getAnnotation(JsonTypeInfo.class);
    if (jsonTypeInfo == null || jsonTypeInfo.property().isBlank()) {
      throw new IllegalStateException(
          "Catalog coverage requires %s to declare a non-blank @JsonTypeInfo property"
              .formatted(sealedType.getName()));
    }
    return jsonTypeInfo.property();
  }

  @SuppressWarnings("unchecked")
  private static String canonicalTypeId(Class<? extends Record> recordType, String suppliedId) {
    if (WorkbookPlan.WorkbookSource.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookSourceTypeName(
              (Class<? extends WorkbookPlan.WorkbookSource>) recordType),
          recordType);
    }
    if (WorkbookPlan.WorkbookPersistence.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookPersistenceTypeName(
              (Class<? extends WorkbookPlan.WorkbookPersistence>) recordType),
          recordType);
    }
    if (WorkbookStep.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookStepTypeName(
              (Class<? extends WorkbookStep>) recordType),
          recordType);
    }
    if (MutationAction.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.mutationActionTypeName(
              (Class<? extends MutationAction>) recordType),
          recordType);
    }
    if (Assertion.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.assertionTypeName((Class<? extends Assertion>) recordType),
          recordType);
    }
    if (InspectionQuery.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.inspectionQueryTypeName(
              (Class<? extends InspectionQuery>) recordType),
          recordType);
    }
    return suppliedId;
  }

  static String requireMatchingCatalogId(
      String suppliedId, String canonicalId, Class<? extends Record> recordType) {
    if (!suppliedId.equals(canonicalId)) {
      throw new IllegalStateException(
          "Catalog type id mismatch for "
              + recordType.getName()
              + ": supplied="
              + suppliedId
              + ", canonical="
              + canonicalId);
    }
    return canonicalId;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static List<FieldEntry> fieldEntries(
      Class<? extends Record> recordType,
      List<String> optionalFields,
      List<CatalogProjectedField> projectedFields) {
    requiredFields(recordType, optionalFields, projectedFields);
    Set<String> optionalFieldSet =
        Stream.concat(
                optionalFields.stream(), projectedFields.stream().map(CatalogProjectedField::name))
            .collect(Collectors.toUnmodifiableSet());
    Map<String, List<String>> projectedFieldsByName = new LinkedHashMap<>();
    for (CatalogProjectedField projectedField : projectedFields) {
      projectedFieldsByName.put(projectedField.name(), projectedField.projectedByFacets());
    }
    return Arrays.stream(recordType.getRecordComponents())
        .filter(CatalogTypeEntryFactory::isCatalogVisible)
        .gather(CatalogGatherers.expandFieldsWithMetadata(optionalFieldSet, projectedFieldsByName))
        .toList();
  }

  private static List<String> recordFields(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents())
        .filter(CatalogTypeEntryFactory::isCatalogVisible)
        .map(RecordComponent::getName)
        .toList();
  }

  private static boolean isCatalogVisible(RecordComponent component) {
    return !component.isAnnotationPresent(CatalogIgnored.class)
        && !component.getAccessor().isAnnotationPresent(CatalogIgnored.class)
        && !component.getAccessor().isAnnotationPresent(JsonIgnore.class);
  }
}
