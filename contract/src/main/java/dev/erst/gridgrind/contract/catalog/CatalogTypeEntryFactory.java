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
import java.util.List;
import java.util.Optional;
import java.util.Set;

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

  static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new CatalogTypeDescriptor(recordType, id, summary, List.of(optionalFields));
  }

  static TypeEntry typeEntry(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    Optional<WorkbookStepTargeting.TargetSurface> targetSurface =
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(recordType);
    return new TypeEntry(
        canonicalTypeId(recordType, id),
        summary,
        fieldEntries(recordType, optionalFields),
        TypeEntryTargetingSupport.targetSelectorEntries(targetSurface),
        targetSurface.flatMap(WorkbookStepTargeting.TargetSurface::rule));
  }

  static List<String> requiredFields(
      Class<? extends Record> recordType, List<String> optionalFields) {
    List<String> recordFields = recordFields(recordType);
    for (String optionalField : optionalFields) {
      if (!recordFields.contains(optionalField)) {
        throw new IllegalStateException(
            "Catalog optional field '%s' does not exist on %s"
                .formatted(optionalField, recordType.getName()));
      }
    }
    return recordFields.stream().filter(field -> !optionalFields.contains(field)).toList();
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

  private static List<FieldEntry> fieldEntries(
      Class<? extends Record> recordType, List<String> optionalFields) {
    requiredFields(recordType, optionalFields);
    Set<String> optionalFieldSet = Set.copyOf(optionalFields);
    return Arrays.stream(recordType.getRecordComponents())
        .filter(CatalogTypeEntryFactory::isCatalogVisible)
        .gather(CatalogGatherers.expandFieldsWithMetadata(optionalFieldSet))
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
