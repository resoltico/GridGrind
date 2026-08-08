package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.CellGridInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellReport;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.CellValueReport;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import dev.erst.gridgrind.excel.foundation.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.excel.foundation.ExcelStoredCellErrorLiteral;
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
              new ComponentKey(CellInput.ErrorValue.class, "error"),
              STORED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellScalarValue.ErrorValue.class, "error"),
              REPORTED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellRowInput.ErrorValues.class, "cells"),
              STORED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellGridInput.ErrorRows.class, "cells"),
              STORED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellValueReport.ErrorValue.class, "errorValue"),
              REPORTED_CELL_ERROR_LITERAL_METADATA),
          Map.entry(
              new ComponentKey(CellReport.ErrorReport.class, "errorValue"),
              REPORTED_CELL_ERROR_LITERAL_METADATA));

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
    return new FieldEntry(
        fieldName,
        optionalFields.contains(fieldName) ? FieldRequirement.OPTIONAL : FieldRequirement.REQUIRED,
        fieldShape(component.getGenericType()),
        metadataOverride.overrideEnumValues(enumValues(component.getGenericType())),
        protocolField == null ? Optional.empty() : protocolField.booleanDefault().value(),
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
    return CatalogFieldShapeRegistry.groupedFieldShape(classType);
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
