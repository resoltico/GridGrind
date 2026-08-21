package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Canonical request-contract support shared by intake, discovery, and diagnostics. */
public final class GridGrindProtocolContractSupport {
  private static final java.util.Set<Class<? extends Record>> REQUEST_RECORD_TYPES =
      discoverRequestRecordTypes();

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
    return effectiveObjectContract(recordType).requiredFields();
  }

  /** Returns the record-owned fields that may be omitted from one protocol shape. */
  public static List<String> optionalFieldNames(Class<? extends Record> recordType) {
    return effectiveObjectContract(recordType).optionalFields();
  }

  /**
   * Returns the effective request-object contract declared by one record's JSON creator.
   *
   * <p>When a record supplies one {@link com.fasterxml.jackson.annotation.JsonCreator}, its named
   * parameters own the wire shape; otherwise the visible record components own it. A creator that
   * disagrees with its record is a program defect, not an ambiguity for request validation to guess
   * through.
   */
  public static EffectiveObjectContract effectiveObjectContract(
      Class<? extends Record> recordType) {
    Objects.requireNonNull(recordType, "recordType must not be null");
    return RequestCreatorContractSupport.effectiveObjectContract(recordType);
  }

  /** Returns the effective selected-union contract, including its required discriminator. */
  public static EffectiveObjectContract effectiveObjectContract(
      Class<? extends Record> recordType, String discriminator) {
    Objects.requireNonNull(discriminator, "discriminator must not be null");
    return RequestCreatorContractSupport.effectiveObjectContract(recordType, discriminator);
  }

  /** Returns the minimal effective contract before a discriminator selects its record subtype. */
  public static EffectiveObjectContract discriminatorContract(String discriminator) {
    Objects.requireNonNull(discriminator, "discriminator must not be null");
    return RequestCreatorContractSupport.discriminatorContract(discriminator);
  }

  /** Returns the JSON property name owned by one record component. */
  public static String wireFieldName(RecordComponent component) {
    Objects.requireNonNull(component, "component must not be null");
    JsonProperty annotation = component.getAccessor().getAnnotation(JsonProperty.class);
    return annotation == null || annotation.value().isBlank()
        ? component.getName()
        : annotation.value();
  }

  /** Returns whether one record is reachable from the authored request graph. */
  public static boolean isRequestInputRecord(Class<? extends Record> recordType) {
    Objects.requireNonNull(recordType, "recordType must not be null");
    return REQUEST_RECORD_TYPES.contains(recordType);
  }

  static java.util.Set<Class<? extends Record>> requestInputRecordTypes() {
    return REQUEST_RECORD_TYPES;
  }

  /**
   * Returns every primitive field reachable from the request graph and its effective requirement.
   *
   * <p>Request intake validates these fields on the raw syntax tree before the typed mapper sees a
   * fragment. This inventory lets that boundary be regression-tested without treating normalized
   * runtime record components as direct JSON bindings.
   */
  public static List<RequestPrimitiveField> requestPrimitiveFields() {
    return requestPrimitiveFields(REQUEST_RECORD_TYPES);
  }

  static List<RequestPrimitiveField> requestPrimitiveFields(
      Set<Class<? extends Record>> recordTypes) {
    Objects.requireNonNull(recordTypes, "recordTypes must not be null");
    return recordTypes.stream()
        .sorted(java.util.Comparator.comparing(Class::getName))
        .flatMap(
            recordType -> {
              EffectiveObjectContract contract = effectiveObjectContract(recordType);
              Set<String> requiredFields = Set.copyOf(contract.requiredFields());
              List<RecordComponent> visibleComponents =
                  RequestCreatorContractSupport.visibleRecordComponents(recordType);
              return contract.fields().stream()
                  .map(
                      fieldName ->
                          Objects.requireNonNull(
                              visibleComponents.stream()
                                  .filter(component -> wireFieldName(component).equals(fieldName))
                                  .findFirst()
                                  .orElse(null),
                              "Effective request contract field has no visible record component"))
                  .filter(component -> component.getType().isPrimitive())
                  .map(
                      component -> {
                        String fieldName = wireFieldName(component);
                        return new RequestPrimitiveField(
                            recordType,
                            fieldName,
                            requiredFields.contains(fieldName),
                            booleanDefault(component));
                      });
            })
        .toList();
  }

  private static Optional<Boolean> booleanDefault(RecordComponent component) {
    ProtocolField field = component.getAnnotation(ProtocolField.class);
    return field == null ? Optional.empty() : field.booleanDefault().value();
  }

  private static java.util.Set<Class<? extends Record>> discoverRequestRecordTypes() {
    java.util.Set<Class<? extends Record>> records = new LinkedHashSet<>();
    collectRequestTypes(WorkbookPlan.class, records, new java.util.HashSet<>());
    return java.util.Set.copyOf(records);
  }

  static void collectRequestTypes(
      Type type, java.util.Set<Class<? extends Record>> records, java.util.Set<Type> visited) {
    if (!visited.add(type)) {
      return;
    }
    if (type instanceof ParameterizedType parameterizedType) {
      for (Type argument : parameterizedType.getActualTypeArguments()) {
        collectRequestTypes(argument, records, visited);
      }
      return;
    }
    if (!(type instanceof Class<?> classType)) {
      return;
    }
    if (classType.isRecord()) {
      @SuppressWarnings("unchecked")
      Class<? extends Record> recordType = (Class<? extends Record>) classType;
      records.add(recordType);
      for (RecordComponent component : recordType.getRecordComponents()) {
        collectRequestTypes(component.getGenericType(), records, visited);
      }
    }
    if (classType.isSealed()) {
      for (Class<?> subtype : classType.getPermittedSubclasses()) {
        collectRequestTypes(subtype, records, visited);
      }
    }
  }

  static Optional<Class<?>> creatorParameterType(
      Class<? extends Record> recordType, String fieldName) {
    return RequestCreatorContractSupport.creatorParameterType(recordType, fieldName);
  }

  /** The complete non-overlapping wire-field contract for one request object. */
  public record EffectiveObjectContract(List<String> requiredFields, List<String> optionalFields) {
    public EffectiveObjectContract {
      requiredFields =
          List.copyOf(Objects.requireNonNull(requiredFields, "requiredFields must not be null"));
      optionalFields =
          List.copyOf(Objects.requireNonNull(optionalFields, "optionalFields must not be null"));
      if (requiredFields.stream().anyMatch(String::isBlank)
          || optionalFields.stream().anyMatch(String::isBlank)) {
        throw new IllegalArgumentException("contract field names must not be blank");
      }
      Set<String> fields = new LinkedHashSet<>(requiredFields);
      if (fields.size() != requiredFields.size()
          || !java.util.Collections.disjoint(fields, optionalFields)) {
        throw new IllegalArgumentException("contract fields must be unique and non-overlapping");
      }
      fields.addAll(optionalFields);
      if (fields.size() != requiredFields.size() + optionalFields.size()) {
        throw new IllegalArgumentException("contract fields must be unique and non-overlapping");
      }
    }

    /** Returns every field accepted by this exact object contract in deterministic order. */
    public List<String> fields() {
      return java.util.stream.Stream.concat(requiredFields.stream(), optionalFields.stream())
          .toList();
    }
  }

  /** One primitive field in the raw request contract. */
  public record RequestPrimitiveField(
      Class<? extends Record> recordType,
      String fieldName,
      boolean required,
      Optional<Boolean> defaultBoolean) {
    public RequestPrimitiveField {
      Objects.requireNonNull(recordType, "recordType must not be null");
      Objects.requireNonNull(fieldName, "fieldName must not be null");
      Objects.requireNonNull(defaultBoolean, "defaultBoolean must not be null");
      if (fieldName.isBlank()) {
        throw new IllegalArgumentException("fieldName must not be blank");
      }
      if (required && defaultBoolean.isPresent()) {
        throw new IllegalArgumentException(
            "required primitive request fields must not declare defaults");
      }
      if (!required && defaultBoolean.isEmpty()) {
        throw new IllegalArgumentException(
            "optional primitive request fields must declare defaults");
      }
    }
  }
}
