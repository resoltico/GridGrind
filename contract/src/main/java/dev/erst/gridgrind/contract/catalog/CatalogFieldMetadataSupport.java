package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.selector.Selector;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Resolves machine-readable field metadata for protocol-catalog record components. */
public final class CatalogFieldMetadataSupport {
  private static final Set<Class<?>> STRING_FIELD_TYPES =
      Set.of(String.class, java.time.LocalDate.class, java.time.LocalDateTime.class);
  private static final Set<Class<?>> BOOLEAN_FIELD_TYPES = Set.of(boolean.class, Boolean.class);
  private static final Set<Class<?>> NUMERIC_FIELD_TYPES =
      Set.of(
          byte.class,
          short.class,
          int.class,
          long.class,
          float.class,
          double.class,
          Byte.class,
          Short.class,
          Integer.class,
          Long.class,
          Float.class,
          Double.class,
          java.math.BigDecimal.class,
          java.math.BigInteger.class);
  private static final Map<Class<?>, String> NESTED_FIELD_SHAPE_GROUPS = nestedFieldShapeGroups();
  private static final Map<Class<?>, List<String>> NESTED_FIELD_SHAPE_UNIONS =
      nestedFieldShapeUnions();
  private static final Map<Class<?>, String> TOP_LEVEL_FIELD_SHAPE_TYPE_SETS =
      topLevelFieldShapeTypeSets();
  private static final Map<Class<?>, String> PLAIN_FIELD_SHAPE_GROUPS = plainFieldShapeGroups();

  private CatalogFieldMetadataSupport() {}

  /** Returns the catalog field entry derived from one reflected record component. */
  public static FieldEntry fieldEntry(RecordComponent component, Set<String> optionalFields) {
    Objects.requireNonNull(component, "component must not be null");
    Objects.requireNonNull(optionalFields, "optionalFields must not be null");
    return new FieldEntry(
        component.getName(),
        optionalFields.contains(component.getName())
            ? FieldRequirement.OPTIONAL
            : FieldRequirement.REQUIRED,
        fieldShape(component.getGenericType()),
        enumValues(component.getGenericType()));
  }

  /** Returns the machine-readable field shape for one record component type. */
  public static FieldShape fieldShape(Type type) {
    Objects.requireNonNull(type, "type must not be null");
    if (type instanceof ParameterizedType parameterizedType) {
      return fieldShape(parameterizedType);
    }
    if (type instanceof Class<?> classType) {
      return fieldShape(classType);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + type);
  }

  /** Returns the machine-readable field shape for one parameterized record component type. */
  public static FieldShape fieldShape(ParameterizedType parameterizedType) {
    Objects.requireNonNull(parameterizedType, "parameterizedType must not be null");
    Type rawType = parameterizedType.getRawType();
    if (rawType == java.util.Optional.class) {
      return fieldShape(singleTypeArgument(parameterizedType, "Optional"));
    }
    if (rawType == java.util.List.class) {
      return new FieldShape.ListShape(fieldShape(singleTypeArgument(parameterizedType, "List")));
    }
    throw new IllegalStateException(
        "Unsupported parameterized catalog field type: " + parameterizedType);
  }

  /** Returns the machine-readable field shape for one non-parameterized record component type. */
  public static FieldShape fieldShape(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    if (STRING_FIELD_TYPES.contains(classType)) {
      return new FieldShape.Scalar(ScalarType.STRING);
    }
    if (BOOLEAN_FIELD_TYPES.contains(classType)) {
      return new FieldShape.Scalar(ScalarType.BOOLEAN);
    }
    if (isNumericType(classType)) {
      return new FieldShape.Scalar(ScalarType.NUMBER);
    }
    if (classType.isEnum()) {
      return new FieldShape.Scalar(ScalarType.STRING);
    }
    List<String> nestedGroupUnion = NESTED_FIELD_SHAPE_UNIONS.get(classType);
    if (nestedGroupUnion == null) {
      nestedGroupUnion =
          lookupAssignableGroupList(NESTED_FIELD_SHAPE_UNIONS, classType).orElse(null);
    }
    if (nestedGroupUnion != null) {
      return new FieldShape.NestedTypeGroupUnionRef(nestedGroupUnion);
    }
    String topLevelTypeSet = TOP_LEVEL_FIELD_SHAPE_TYPE_SETS.get(classType);
    if (topLevelTypeSet == null) {
      topLevelTypeSet =
          lookupAssignableGroup(TOP_LEVEL_FIELD_SHAPE_TYPE_SETS, classType).orElse(null);
    }
    if (topLevelTypeSet != null) {
      return new FieldShape.TopLevelTypeSetRef(topLevelTypeSet);
    }
    String nestedGroup = NESTED_FIELD_SHAPE_GROUPS.get(classType);
    if (nestedGroup == null) {
      nestedGroup = lookupAssignableGroup(NESTED_FIELD_SHAPE_GROUPS, classType).orElse(null);
    }
    if (nestedGroup != null) {
      return new FieldShape.NestedTypeGroupRef(nestedGroup);
    }
    String plainGroup = PLAIN_FIELD_SHAPE_GROUPS.get(classType);
    if (plainGroup == null) {
      plainGroup = lookupAssignableGroup(PLAIN_FIELD_SHAPE_GROUPS, classType).orElse(null);
    }
    if (plainGroup != null) {
      return new FieldShape.PlainTypeGroupRef(plainGroup);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + classType.getName());
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

  /** Returns whether one non-parameterized record component type is represented as JSON NUMBER. */
  public static boolean isNumericType(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    return NUMERIC_FIELD_TYPES.contains(classType);
  }

  /** Returns the set of all sealed types registered in the nested field-shape group map. */
  public static Set<Class<?>> registeredNestedTypes() {
    return NESTED_FIELD_SHAPE_GROUPS.keySet();
  }

  /** Returns the set of all record types registered in the plain field-shape group map. */
  public static Set<Class<?>> registeredPlainTypes() {
    return PLAIN_FIELD_SHAPE_GROUPS.keySet();
  }

  /** Validates that one nested sealed input type maps to the published field-shape group. */
  public static void validateNestedTypeGroupMapping(Class<?> sealedType, String expectedGroup) {
    Objects.requireNonNull(sealedType, "sealedType must not be null");
    String mappedGroup = NESTED_FIELD_SHAPE_GROUPS.get(sealedType);
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

  /** Validates that one plain record input type maps to the published field-shape group. */
  public static void validatePlainTypeGroupMapping(Class<?> recordType, String expectedGroup) {
    Objects.requireNonNull(recordType, "recordType must not be null");
    String mappedGroup = PLAIN_FIELD_SHAPE_GROUPS.get(recordType);
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

  private static java.util.List<String> enumValues(Type type) {
    if (type instanceof ParameterizedType parameterizedType
        && parameterizedType.getRawType() == java.util.Optional.class) {
      return enumValues(singleTypeArgument(parameterizedType, "Optional"));
    }
    if (type instanceof Class<?> classType && classType.isEnum()) {
      return Arrays.stream(classType.getEnumConstants())
          .map(value -> ((Enum<?>) value).name())
          .toList();
    }
    return java.util.List.of();
  }

  private static Type singleTypeArgument(ParameterizedType parameterizedType, String typeName) {
    Type[] typeArguments = parameterizedType.getActualTypeArguments();
    if (typeArguments.length != 1) {
      throw new IllegalStateException(
          typeName + " field must declare exactly one type argument: " + parameterizedType);
    }
    return typeArguments[0];
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
