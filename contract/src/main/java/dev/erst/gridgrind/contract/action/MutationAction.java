package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.List;
import java.util.Objects;

/** One validated mutation action expressed in protocol form. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
public sealed interface MutationAction
    permits WorkbookMutationAction,
        CellMutationAction,
        DrawingMutationAction,
        StructuredMutationAction {
  /** Returns the SCREAMING_SNAKE_CASE type name of this action as used in the wire protocol. */
  default String actionType() {
    return GridGrindProtocolTypeNames.mutationActionTypeName(
        getClass().asSubclass(MutationAction.class));
  }

  /** Returns the selector types accepted by one concrete mutation action instance. */
  static Class<? extends Selector>[] allowedTargetTypes(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    return allowedTargetTypesForType(action.getClass().asSubclass(MutationAction.class));
  }

  /** Returns the selector types accepted by one concrete mutation action type. */
  static Class<? extends Selector>[] allowedTargetTypesForType(
      Class<? extends MutationAction> actionType) {
    Objects.requireNonNull(actionType, "actionType must not be null");
    if (!actionType.isRecord()) {
      throw new IllegalArgumentException(
          "No target-type mapping configured for action class " + actionType.getName());
    }
    try {
      return ProtocolTypeMetadataSupport.staticTargetSelectors(actionType);
    } catch (IllegalStateException exception) {
      throw new IllegalArgumentException(
          "No target-type mapping configured for action class " + actionType.getName(), exception);
    }
  }

  /** Shared validation helpers for MutationAction compact constructors. */
  final class Validation {
    private Validation() {}

    static void requireNonBlank(String value, String fieldName) {
      MutationActionNameValidation.requireNonBlank(value, fieldName);
    }

    static void requireSheetName(String value, String fieldName) { // LIM-003
      MutationActionNameValidation.requireSheetName(value, fieldName);
    }

    static void requireNonNegative(int value, String fieldName) {
      MutationActionNumericValidation.requireNonNegative(value, fieldName);
    }

    static void requirePositive(int value, String fieldName) {
      MutationActionNumericValidation.requirePositive(value, fieldName);
    }

    static void requireNonZero(int value, String fieldName) {
      MutationActionNumericValidation.requireNonZero(value, fieldName);
    }

    static void requireRowIndex(int value, String fieldName) {
      MutationActionNumericValidation.requireRowIndex(value, fieldName);
    }

    static void requireColumnIndex(int value, String fieldName) {
      MutationActionNumericValidation.requireColumnIndex(value, fieldName);
    }

    static void requireOrderedSpan(
        int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
      MutationActionNumericValidation.requireOrderedSpan(
          firstValue, lastValue, firstFieldName, lastFieldName);
    }

    static void requireColumnWidthCharacters(double widthCharacters) { // LIM-004
      MutationActionNumericValidation.requireColumnWidthCharacters(widthCharacters);
    }

    static void requireRowHeightPoints(double heightPoints) { // LIM-005
      MutationActionNumericValidation.requireRowHeightPoints(heightPoints);
    }

    static void requireNamedRangeName(String name) {
      MutationActionNameValidation.requireNamedRangeName(name);
    }

    static void requirePivotTableName(String name) {
      MutationActionNameValidation.requirePivotTableName(name);
    }

    static void requireZoomPercent(int zoomPercent) { // LIM-022
      MutationActionNumericValidation.requireZoomPercent(zoomPercent);
    }

    static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
      return MutationActionCollectionValidation.copyRows(rows);
    }

    static List<List<CellInput>> freezeRows(List<List<CellInput>> rows) {
      return MutationActionCollectionValidation.freezeRows(rows);
    }

    static List<String> copySheetNames(List<String> sheetNames, String fieldName) {
      return MutationActionCollectionValidation.copySheetNames(sheetNames, fieldName);
    }

    static void requireDistinct(List<String> values, String fieldName) {
      MutationActionCollectionValidation.requireDistinct(values, fieldName);
    }

    static void requireRectangularRows(List<List<CellInput>> rows) {
      MutationActionCollectionValidation.requireRectangularRows(rows);
    }
  }
}
