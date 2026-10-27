package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.semantics.WorkbookOperationContracts;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;
import java.util.Objects;

/** One validated mutation action expressed in protocol form. */
@Union(discriminator = "type")
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_COLUMN_WIDTH,
  dev.erst.gridgrind.core.limits.Limit.ID_ROW_HEIGHT,
  dev.erst.gridgrind.core.limits.Limit.ID_SHEET_ZOOM_PERCENTAGE
})
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
    return WorkbookOperationContracts.targetSelectorsFor(action);
  }

  /** Returns the selector types accepted by one concrete mutation action type. */
  static Class<? extends Selector>[] allowedTargetTypesForType(
      Class<? extends MutationAction> actionType) {
    Objects.requireNonNull(actionType, "actionType must not be null");
    if (!actionType.isRecord()) {
      throw noTargetTypeMapping(actionType);
    }
    try {
      return WorkbookOperationContracts.staticTargetSelectorsFor(actionType);
    } catch (IllegalArgumentException | IllegalStateException exception) {
      throw noTargetTypeMapping(actionType, exception);
    }
  }

  private static IllegalArgumentException noTargetTypeMapping(
      Class<? extends MutationAction> actionType) {
    return new IllegalArgumentException(
        "No target-type mapping configured for action class " + actionType.getName());
  }

  private static IllegalArgumentException noTargetTypeMapping(
      Class<? extends MutationAction> actionType, RuntimeException cause) {
    return new IllegalArgumentException(
        "No target-type mapping configured for action class " + actionType.getName(), cause);
  }

  /** Shared validation helpers for MutationAction compact constructors. */
  final class Validation {
    private Validation() {}

    static void requireNonBlank(String value, String fieldName) {
      MutationActionNameValidation.requireNonBlank(value, fieldName);
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

    static void requireColumnWidthCharacters(double widthCharacters) {
      MutationActionNumericValidation.requireColumnWidthCharacters(widthCharacters);
    }

    static void requireRowHeightPoints(double heightPoints) {
      MutationActionNumericValidation.requireRowHeightPoints(heightPoints);
    }

    static void requirePivotTableName(String name) {
      MutationActionNameValidation.requirePivotTableName(name);
    }

    static void requireZoomPercent(int zoomPercent) {
      MutationActionNumericValidation.requireZoomPercent(zoomPercent);
    }

    static List<List<CellValueSpec>> copyRows(List<List<CellValueSpec>> rows) {
      return MutationActionCollectionValidation.copyRows(rows);
    }

    static List<List<CellValueSpec>> freezeRows(List<List<CellValueSpec>> rows) {
      return MutationActionCollectionValidation.freezeRows(rows);
    }

    static List<SheetName> copySheetNames(List<SheetName> sheetNames, String fieldName) {
      return MutationActionCollectionValidation.copySheetNames(sheetNames, fieldName);
    }

    static void requireDistinct(List<String> values, String fieldName) {
      MutationActionCollectionValidation.requireDistinct(values, fieldName);
    }

    static void requireRectangularRows(List<List<CellValueSpec>> rows) {
      MutationActionCollectionValidation.requireRectangularRows(rows);
    }
  }
}
