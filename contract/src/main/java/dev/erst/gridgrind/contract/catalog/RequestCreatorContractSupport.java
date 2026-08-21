package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.dto.ProtocolBooleanDefault;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import java.lang.reflect.Executable;
import java.lang.reflect.Method;
import java.lang.reflect.Parameter;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

/** Derives the effective JSON-creator contract for request records. */
final class RequestCreatorContractSupport {
  private RequestCreatorContractSupport() {}

  static GridGrindProtocolContractSupport.EffectiveObjectContract effectiveObjectContract(
      Class<? extends Record> recordType) {
    List<RecordComponent> components = visibleRecordComponents(recordType);
    Map<String, RecordComponent> componentsByName =
        components.stream()
            .collect(
                Collectors.toMap(
                    GridGrindProtocolContractSupport::wireFieldName,
                    component -> component,
                    (left, right) -> {
                      throw new IllegalStateException(
                          "Record exposes duplicate JSON field names: " + recordType.getName());
                    },
                    java.util.LinkedHashMap::new));
    List<String> fields =
        creatorFieldNames(recordType).orElseGet(() -> List.copyOf(componentsByName.keySet()));
    if (!componentsByName.keySet().equals(new LinkedHashSet<>(fields))) {
      throw new IllegalStateException(
          "JSON creator fields must exactly match visible record fields for "
              + recordType.getName()
              + ": creator="
              + fields
              + ", record="
              + componentsByName.keySet());
    }
    List<String> required =
        fields.stream()
            .filter(field -> !isOptional(componentFor(componentsByName, field)))
            .toList();
    List<String> optional =
        fields.stream().filter(field -> isOptional(componentFor(componentsByName, field))).toList();
    components.forEach(component -> validatePrimitiveDefault(component, isOptional(component)));
    return new GridGrindProtocolContractSupport.EffectiveObjectContract(required, optional);
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

  static Optional<Class<?>> creatorParameterType(
      Class<? extends Record> recordType, String fieldName) {
    List<Executable> creators = jsonCreators(recordType);
    if (creators.size() != 1) {
      return Optional.empty();
    }
    return Arrays.stream(creators.getFirst().getParameters())
        .filter(parameter -> fieldName.equals(jsonPropertyName(parameter)))
        .<Class<?>>map(Parameter::getType)
        .findFirst();
  }

  static List<RecordComponent> visibleRecordComponents(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents())
        .filter(component -> !component.getAccessor().isAnnotationPresent(JsonIgnore.class))
        .toList();
  }

  private static RecordComponent componentFor(
      Map<String, RecordComponent> componentsByName, String field) {
    // The exact field-set check above establishes that every creator field has a component.
    return java.util.Objects.requireNonNull(componentsByName.get(field));
  }

  private static Optional<List<String>> creatorFieldNames(Class<? extends Record> recordType) {
    List<Executable> creators = jsonCreators(recordType);
    if (creators.isEmpty()) {
      return Optional.empty();
    }
    if (creators.size() != 1) {
      throw new IllegalStateException(
          "Request record must declare at most one @JsonCreator: " + recordType.getName());
    }
    List<String> fields =
        Arrays.stream(creators.getFirst().getParameters())
            .map(
                parameter -> {
                  String fieldName = jsonPropertyName(parameter);
                  if (fieldName.isBlank()) {
                    throw new IllegalStateException(
                        "Every @JsonCreator parameter must declare @JsonProperty: "
                            + recordType.getName());
                  }
                  return fieldName;
                })
            .toList();
    if (fields.size() != new LinkedHashSet<>(fields).size()) {
      throw new IllegalStateException(
          "JSON creator must not declare duplicate property names: " + recordType.getName());
    }
    return Optional.of(fields);
  }

  private static List<Executable> jsonCreators(Class<? extends Record> recordType) {
    return java.util.stream.Stream.concat(
            Arrays.stream(recordType.getDeclaredConstructors())
                .filter(constructor -> constructor.isAnnotationPresent(JsonCreator.class)),
            Arrays.stream(recordType.getDeclaredMethods())
                .filter(method -> method.isAnnotationPresent(JsonCreator.class))
                .map(Method.class::cast))
        .map(Executable.class::cast)
        .toList();
  }

  private static String jsonPropertyName(Parameter parameter) {
    JsonProperty property = parameter.getAnnotation(JsonProperty.class);
    return property == null ? "" : property.value();
  }

  private static boolean isOptional(RecordComponent component) {
    ProtocolField field = component.getAnnotation(ProtocolField.class);
    return component.getType() == Optional.class
        || isNonAbsent(component)
        || (field != null && field.optional());
  }

  private static void validatePrimitiveDefault(RecordComponent component, boolean optional) {
    ProtocolField field = component.getAnnotation(ProtocolField.class);
    ProtocolBooleanDefault booleanDefault =
        field == null ? ProtocolBooleanDefault.UNSPECIFIED : field.booleanDefault();
    if (booleanDefault == ProtocolBooleanDefault.UNSPECIFIED) {
      rejectUndeclaredOptionalPrimitive(component, optional);
      return;
    }
    rejectInapplicableBooleanDefault(component, optional);
  }

  private static void rejectInapplicableBooleanDefault(
      RecordComponent component, boolean optional) {
    if (component.getType() != boolean.class) {
      throw new IllegalStateException(
          "Only boolean request fields may declare a boolean default: "
              + component.getDeclaringRecord().getName()
              + "."
              + GridGrindProtocolContractSupport.wireFieldName(component));
    }
    if (!optional) {
      throw new IllegalStateException(
          "Only optional request fields may declare a default: "
              + component.getDeclaringRecord().getName()
              + "."
              + GridGrindProtocolContractSupport.wireFieldName(component));
    }
  }

  private static void rejectUndeclaredOptionalPrimitive(
      RecordComponent component, boolean optional) {
    if (optional && component.getType().isPrimitive()) {
      throw new IllegalStateException(
          "Optional primitive request fields must declare an explicit default: "
              + component.getDeclaringRecord().getName()
              + "."
              + GridGrindProtocolContractSupport.wireFieldName(component));
    }
  }

  private static boolean isNonAbsent(RecordComponent component) {
    JsonInclude include = component.getAccessor().getAnnotation(JsonInclude.class);
    return include != null && include.value() == JsonInclude.Include.NON_ABSENT;
  }
}
