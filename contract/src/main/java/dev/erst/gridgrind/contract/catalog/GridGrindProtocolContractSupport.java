package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/** Canonical request-contract support shared by intake, discovery, and diagnostics. */
public final class GridGrindProtocolContractSupport {
  private static final Map<Class<? extends Record>, List<String>> REQUIRED_FIELDS_BY_RECORD_TYPE =
      buildRequiredFieldsByRecordType();

  private GridGrindProtocolContractSupport() {}

  /** Returns the discriminator field for one typed protocol root when the type declares one. */
  public static Optional<String> discriminatorField(Class<?> type) {
    Objects.requireNonNull(type, "type must not be null");
    JsonTypeInfo jsonTypeInfo = type.getAnnotation(JsonTypeInfo.class);
    if (jsonTypeInfo == null || jsonTypeInfo.property().isBlank()) {
      return Optional.empty();
    }
    return Optional.of(jsonTypeInfo.property());
  }

  /** Returns the effective required request fields for one record type. */
  public static List<String> requiredFieldNames(Class<? extends Record> recordType) {
    Objects.requireNonNull(recordType, "recordType must not be null");
    List<String> knownRequiredFields = REQUIRED_FIELDS_BY_RECORD_TYPE.get(recordType);
    if (knownRequiredFields != null) {
      return knownRequiredFields;
    }
    return fallbackRequiredFields(recordType);
  }

  private static Map<Class<? extends Record>, List<String>> buildRequiredFieldsByRecordType() {
    Map<Class<? extends Record>, List<String>> requiredFieldsByRecordType =
        new ConcurrentHashMap<>();
    register(
        requiredFieldsByRecordType,
        WorkbookPlan.class,
        GridGrindProtocolCatalog.catalog().requestType().fields().stream()
            .filter(field -> field.requirement() == FieldRequirement.REQUIRED)
            .map(FieldEntry::name)
            .toList());
    for (CatalogTypeDescriptor descriptor : GridGrindProtocolCatalogTypeDescriptors.ALL_TYPES) {
      register(
          requiredFieldsByRecordType,
          descriptor.recordType(),
          CatalogTypeEntryFactory.requiredFields(
              descriptor.recordType(), descriptor.optionalFields()));
    }
    for (CatalogNestedTypeDescriptor group :
        GridGrindProtocolCatalogFieldGroupSupport.NESTED_TYPE_GROUPS) {
      for (CatalogTypeDescriptor descriptor : group.typeDescriptors()) {
        register(
            requiredFieldsByRecordType,
            descriptor.recordType(),
            CatalogTypeEntryFactory.requiredFields(
                descriptor.recordType(), descriptor.optionalFields()));
      }
    }
    for (CatalogPlainTypeDescriptor descriptor :
        GridGrindProtocolCatalogFieldGroupSupport.PLAIN_TYPE_DESCRIPTORS) {
      register(
          requiredFieldsByRecordType,
          descriptor.recordType(),
          CatalogTypeEntryFactory.requiredFields(
              descriptor.recordType(), descriptor.optionalFields()));
    }
    return Map.copyOf(requiredFieldsByRecordType);
  }

  static void register(
      Map<Class<? extends Record>, List<String>> requiredFieldsByRecordType,
      Class<? extends Record> recordType,
      List<String> requiredFields) {
    List<String> previous =
        requiredFieldsByRecordType.putIfAbsent(recordType, List.copyOf(requiredFields));
    if (previous != null && !previous.equals(requiredFields)) {
      throw new IllegalStateException(
          "Conflicting required-field definitions for "
              + recordType.getName()
              + ": "
              + previous
              + " vs "
              + requiredFields);
    }
  }

  private static List<String> fallbackRequiredFields(Class<? extends Record> recordType) {
    Set<String> optionalFields = new LinkedHashSet<>();
    for (RecordComponent component : recordType.getRecordComponents()) {
      if (component.getType() == Optional.class) {
        optionalFields.add(component.getName());
      }
    }
    ProtocolTypeMetadata metadata = recordType.getAnnotation(ProtocolTypeMetadata.class);
    if (metadata != null) {
      optionalFields.addAll(Set.of(metadata.optionalFields()));
    }
    return Arrays.stream(recordType.getRecordComponents())
        .map(RecordComponent::getName)
        .filter(fieldName -> !optionalFields.contains(fieldName))
        .toList();
  }
}
