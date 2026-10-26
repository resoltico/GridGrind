package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.CatalogIgnored;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

/** Derives request field presence solely from canonical record components and metadata. */
final class RequestCreatorContractSupport {
  private RequestCreatorContractSupport() {}

  static GridGrindProtocolContractSupport.EffectiveObjectContract effectiveObjectContract(
      Class<? extends Record> recordType) {
    List<RecordComponent> components = visibleRecordComponents(recordType);
    components.forEach(RequestCreatorContractSupport::validatePrimitiveDefault);
    return new GridGrindProtocolContractSupport.EffectiveObjectContract(
        components.stream()
            .filter(component -> !isOptional(component))
            .map(RecordComponent::getName)
            .toList(),
        components.stream()
            .filter(RequestCreatorContractSupport::isOptional)
            .map(RecordComponent::getName)
            .toList());
  }

  static GridGrindProtocolContractSupport.EffectiveObjectContract effectiveObjectContract(
      Class<? extends Record> recordType, String discriminator) {
    GridGrindProtocolContractSupport.EffectiveObjectContract recordContract =
        effectiveObjectContract(recordType);
    if (recordContract.optionalFields().contains(discriminator)) {
      throw new IllegalStateException(
          "A discriminator cannot be optional in the selected request contract: " + discriminator);
    }
    List<String> required = new java.util.ArrayList<>(recordContract.requiredFields());
    if (!required.contains(discriminator)) {
      required.add(discriminator);
    }
    List<String> optional = recordContract.optionalFields();
    return new GridGrindProtocolContractSupport.EffectiveObjectContract(required, optional);
  }

  static GridGrindProtocolContractSupport.EffectiveObjectContract discriminatorContract(
      String discriminator) {
    if (discriminator.isBlank()) {
      throw new IllegalArgumentException("discriminator must not be blank");
    }
    return new GridGrindProtocolContractSupport.EffectiveObjectContract(
        List.of(discriminator), List.of());
  }

  static List<RecordComponent> visibleRecordComponents(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents())
        .filter(
            component ->
                !component.isAnnotationPresent(CatalogIgnored.class)
                    && !component.getAccessor().isAnnotationPresent(CatalogIgnored.class))
        .toList();
  }

  private static boolean isOptional(RecordComponent component) {
    ProtocolField field = component.getAnnotation(ProtocolField.class);
    return component.getType() == Optional.class
        || component.isAnnotationPresent(WireDefault.class)
        || (field != null && field.optional());
  }

  private static void validatePrimitiveDefault(RecordComponent component) {
    if (isOptional(component)
        && component.getType().isPrimitive()
        && !component.isAnnotationPresent(WireDefault.class)) {
      throw new IllegalStateException(
          "Optional primitive request fields must declare an explicit default: "
              + component.getDeclaringRecord().getName()
              + "."
              + component.getName());
    }
  }
}
