package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.vocabulary.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.core.vocabulary.ExcelStoredCellErrorLiteral;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
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
      Set.of(
          String.class,
          SecretReference.class,
          java.nio.file.Path.class,
          java.time.LocalDate.class,
          java.time.LocalDateTime.class);
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
  private static final FieldMetadataOverride STORED_CELL_ERROR_LITERAL_METADATA =
      new FieldMetadataOverride(
          ExcelStoredCellErrorLiteral.orderedWireValues(),
          CatalogEnumValueDocumentationSupport.storedCellErrorLiteralDocs());
  private static final FieldMetadataOverride REPORTED_CELL_ERROR_LITERAL_METADATA =
      new FieldMetadataOverride(
          ExcelReportedCellErrorLiteral.orderedWireValues(),
          CatalogEnumValueDocumentationSupport.reportedCellErrorLiteralDocs());
  private static final Map<ComponentKey, FieldMetadataOverride> FIELD_METADATA_OVERRIDES =
      Map.ofEntries(
          Map.entry(
              new ComponentKey(CellValueSpec.ErrorValue.class, "value"),
              STORED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellValue.ErrorValue.class, "value"),
              REPORTED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellRowSpec.ErrorValues.class, "cells"),
              STORED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellGridSpec.ErrorRows.class, "cells"),
              STORED_CELL_ERROR_LITERAL_METADATA));

  private CatalogFieldMetadataSupport() {}

  /** Returns the catalog field entry derived from one reflected record component. */
  public static FieldEntry fieldEntry(RecordComponent component, Set<String> optionalFields) {
    return fieldEntry(component, optionalFields, Map.of());
  }

  /** Returns the catalog field entry derived from one reflected record component. */
  public static FieldEntry fieldEntry(
      RecordComponent component,
      Set<String> optionalFields,
      Map<String, List<String>> projectedFieldsByName) {
    Objects.requireNonNull(component, "component must not be null");
    Objects.requireNonNull(optionalFields, "optionalFields must not be null");
    Objects.requireNonNull(projectedFieldsByName, "projectedFieldsByName must not be null");
    FieldMetadataOverride metadataOverride =
        FIELD_METADATA_OVERRIDES.getOrDefault(
            new ComponentKey(component.getDeclaringRecord(), component.getName()),
            FieldMetadataOverride.NONE);
    ProtocolField protocolField = component.getAnnotation(ProtocolField.class);
    String fieldName = GridGrindProtocolContractSupport.wireFieldName(component);
    FieldShape shape =
        applyConstraints(
            fieldShape(component.getGenericType()),
            CatalogFieldConstraints.forComponent(component.getDeclaringRecord(), fieldName));
    return new FieldEntry(
        fieldName,
        optionalFields.contains(fieldName) ? FieldRequirement.OPTIONAL : FieldRequirement.REQUIRED,
        shape,
        metadataOverride.overrideEnumValues(enumValues(component.getGenericType())),
        dev.erst.gridgrind.core.metadata.ModelDefaults.publishedValue(component),
        Optional.ofNullable(
                component.getAnnotation(dev.erst.gridgrind.core.metadata.WireDefault.class))
            .map(dev.erst.gridgrind.core.metadata.WireDefault::from)
            .filter(value -> !value.isEmpty()),
        metadataOverride.overrideEnumValueDocs(enumValueDocs(component.getGenericType())),
        projectedFieldsByName.getOrDefault(fieldName, List.of()),
        protocolField != null && protocolField.secret());
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
    if (WireScalar.isValueType(classType)) {
      RecordComponent value = classType.getRecordComponents()[0];
      return applyConstraints(
          fieldShape(value.getGenericType()),
          CatalogFieldConstraints.forComponent(classType, value.getName()));
    }
    return scalarFieldShape(classType)
        .or(() -> groupedFieldShape(classType))
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Unsupported catalog field type: " + classType.getName()));
  }

  /** Returns whether one non-parameterized record component type is represented as JSON NUMBER. */
  public static boolean isNumericType(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    return NUMERIC_FIELD_TYPES.contains(classType);
  }

  private static Optional<FieldShape> scalarFieldShape(Class<?> classType) {
    if (STRING_FIELD_TYPES.contains(classType) || classType.isEnum()) {
      return Optional.of(new FieldShape.Scalar(ScalarType.STRING));
    }
    if (BOOLEAN_FIELD_TYPES.contains(classType)) {
      return Optional.of(new FieldShape.Scalar(ScalarType.BOOLEAN));
    }
    if (isNumericType(classType)) {
      return Optional.of(new FieldShape.Scalar(ScalarType.NUMBER));
    }
    return Optional.empty();
  }

  private static Optional<FieldShape> groupedFieldShape(Class<?> classType) {
    return CatalogGroups.fieldShape(classType);
  }

  private static FieldShape applyConstraints(FieldShape shape, List<FieldConstraint> constraints) {
    if (constraints.isEmpty()) {
      return shape;
    }
    FieldShape.Scalar scalar = (FieldShape.Scalar) shape;
    return new FieldShape.Scalar(scalar.scalarType(), constraints);
  }

  static java.util.List<String> enumValues(Type type) {
    if (type instanceof ParameterizedType parameterizedType
        && parameterizedType.getRawType() == java.util.Optional.class) {
      return enumValues(singleTypeArgument(parameterizedType, "Optional"));
    }
    if (type instanceof ParameterizedType parameterizedType
        && parameterizedType.getRawType() == java.util.List.class) {
      return enumValues(singleTypeArgument(parameterizedType, "List"));
    }
    if (type == ExcelStoredCellErrorLiteral.class) {
      return ExcelStoredCellErrorLiteral.orderedWireValues();
    }
    if (type == ExcelReportedCellErrorLiteral.class) {
      return ExcelReportedCellErrorLiteral.orderedWireValues();
    }
    if (type instanceof Class<?> classType && classType.isEnum()) {
      return Arrays.stream(classType.getEnumConstants())
          .map(value -> ((Enum<?>) value).name())
          .toList();
    }
    return java.util.List.of();
  }

  static List<EnumValueDocEntry> enumValueDocs(Type type) {
    return CatalogEnumValueDocumentationSupport.enumValueDocs(type);
  }

  private record ComponentKey(Class<?> declaringRecord, String componentName) {}

  private record FieldMetadataOverride(
      List<String> enumValues, List<EnumValueDocEntry> enumValueDocs) {
    private static final FieldMetadataOverride NONE =
        new FieldMetadataOverride(List.of(), List.of());

    private List<String> overrideEnumValues(List<String> discoveredValues) {
      return enumValues.isEmpty() ? discoveredValues : enumValues;
    }

    private List<EnumValueDocEntry> overrideEnumValueDocs(List<EnumValueDocEntry> discoveredDocs) {
      return enumValueDocs.isEmpty() ? discoveredDocs : enumValueDocs;
    }
  }

  static Type singleTypeArgument(ParameterizedType parameterizedType, String typeName) {
    Type[] typeArguments = parameterizedType.getActualTypeArguments();
    if (typeArguments.length != 1) {
      throw new IllegalStateException(
          typeName + " field must declare exactly one type argument: " + parameterizedType);
    }
    return typeArguments[0];
  }
}
