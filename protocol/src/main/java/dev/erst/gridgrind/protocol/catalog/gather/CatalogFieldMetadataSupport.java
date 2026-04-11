package dev.erst.gridgrind.protocol.catalog.gather;

import dev.erst.gridgrind.protocol.catalog.FieldEntry;
import dev.erst.gridgrind.protocol.catalog.FieldRequirement;
import dev.erst.gridgrind.protocol.catalog.FieldShape;
import dev.erst.gridgrind.protocol.catalog.ScalarType;
import dev.erst.gridgrind.protocol.dto.*;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
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
  private static final Map<Class<?>, String> NESTED_FIELD_SHAPE_GROUPS =
      Map.ofEntries(
          Map.entry(CellInput.class, "cellInputTypes"),
          Map.entry(HyperlinkTarget.class, "hyperlinkTargetTypes"),
          Map.entry(PaneInput.class, "paneTypes"),
          Map.entry(SheetCopyPosition.class, "sheetCopyPositionTypes"),
          Map.entry(RowSpanInput.class, "rowSpanTypes"),
          Map.entry(ColumnSpanInput.class, "columnSpanTypes"),
          Map.entry(CellSelection.class, "cellSelectionTypes"),
          Map.entry(RangeSelection.class, "rangeSelectionTypes"),
          Map.entry(SheetSelection.class, "sheetSelectionTypes"),
          Map.entry(TableSelection.class, "tableSelectionTypes"),
          Map.entry(NamedRangeSelection.class, "namedRangeSelectionTypes"),
          Map.entry(NamedRangeScope.class, "namedRangeScopeTypes"),
          Map.entry(NamedRangeSelector.class, "namedRangeSelectorTypes"),
          Map.entry(DataValidationRuleInput.class, "dataValidationRuleTypes"),
          Map.entry(AutofilterFilterCriterionInput.class, "autofilterFilterCriterionTypes"),
          Map.entry(ConditionalFormattingRuleInput.class, "conditionalFormattingRuleTypes"),
          Map.entry(PrintAreaInput.class, "printAreaTypes"),
          Map.entry(PrintScalingInput.class, "printScalingTypes"),
          Map.entry(PrintTitleRowsInput.class, "printTitleRowsTypes"),
          Map.entry(PrintTitleColumnsInput.class, "printTitleColumnsTypes"),
          Map.entry(TableStyleInput.class, "tableStyleTypes"),
          Map.entry(FontHeightInput.class, "fontHeightTypes"));
  private static final Map<Class<?>, String> PLAIN_FIELD_SHAPE_GROUPS =
      Map.ofEntries(
          Map.entry(CommentInput.class, "commentInputType"),
          Map.entry(CommentAnchorInput.class, "commentAnchorInputType"),
          Map.entry(NamedRangeTarget.class, "namedRangeTargetType"),
          Map.entry(SheetProtectionSettings.class, "sheetProtectionSettingsType"),
          Map.entry(CellStyleInput.class, "cellStyleInputType"),
          Map.entry(CellAlignmentInput.class, "cellAlignmentInputType"),
          Map.entry(CellFontInput.class, "cellFontInputType"),
          Map.entry(ColorInput.class, "colorInputType"),
          Map.entry(RichTextRunInput.class, "richTextRunInputType"),
          Map.entry(CellFillInput.class, "cellFillInputType"),
          Map.entry(CellGradientFillInput.class, "cellGradientFillInputType"),
          Map.entry(CellGradientStopInput.class, "cellGradientStopInputType"),
          Map.entry(CellBorderInput.class, "cellBorderInputType"),
          Map.entry(CellBorderSideInput.class, "cellBorderSideInputType"),
          Map.entry(CellProtectionInput.class, "cellProtectionInputType"),
          Map.entry(DataValidationInput.class, "dataValidationInputType"),
          Map.entry(DataValidationPromptInput.class, "dataValidationPromptInputType"),
          Map.entry(DataValidationErrorAlertInput.class, "dataValidationErrorAlertInputType"),
          Map.entry(
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "autofilterCustomConditionInputType"),
          Map.entry(AutofilterFilterColumnInput.class, "autofilterFilterColumnInputType"),
          Map.entry(AutofilterSortConditionInput.class, "autofilterSortConditionInputType"),
          Map.entry(AutofilterSortStateInput.class, "autofilterSortStateInputType"),
          Map.entry(ConditionalFormattingBlockInput.class, "conditionalFormattingBlockInputType"),
          Map.entry(
              ConditionalFormattingThresholdInput.class, "conditionalFormattingThresholdInputType"),
          Map.entry(HeaderFooterTextInput.class, "headerFooterTextInputType"),
          Map.entry(DifferentialStyleInput.class, "differentialStyleInputType"),
          Map.entry(DifferentialBorderInput.class, "differentialBorderInputType"),
          Map.entry(DifferentialBorderSideInput.class, "differentialBorderSideInputType"),
          Map.entry(PrintLayoutInput.class, "printLayoutInputType"),
          Map.entry(PrintMarginsInput.class, "printMarginsInputType"),
          Map.entry(PrintSetupInput.class, "printSetupInputType"),
          Map.entry(TableColumnInput.class, "tableColumnInputType"),
          Map.entry(TableInput.class, "tableInputType"),
          Map.entry(WorkbookProtectionInput.class, "workbookProtectionInputType"));

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
    if (rawType == java.util.List.class) {
      Type[] typeArguments = parameterizedType.getActualTypeArguments();
      if (typeArguments.length != 1) {
        throw new IllegalStateException(
            "List field must declare exactly one type argument: " + parameterizedType);
      }
      return new FieldShape.ListShape(fieldShape(typeArguments[0]));
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
    String nestedGroup = NESTED_FIELD_SHAPE_GROUPS.get(classType);
    if (nestedGroup != null) {
      return new FieldShape.NestedTypeGroupRef(nestedGroup);
    }
    String plainGroup = PLAIN_FIELD_SHAPE_GROUPS.get(classType);
    if (plainGroup != null) {
      return new FieldShape.PlainTypeGroupRef(plainGroup);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + classType.getName());
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
    if (type instanceof Class<?> classType && classType.isEnum()) {
      return Arrays.stream(classType.getEnumConstants())
          .map(value -> ((Enum<?>) value).name())
          .toList();
    }
    return java.util.List.of();
  }
}
