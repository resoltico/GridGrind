package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.ArrayList;
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
      Objects.requireNonNull(value, fieldName + " must not be null");
      if (value.isBlank()) {
        throw new IllegalArgumentException(fieldName + " must not be blank");
      }
    }

    static void requireSheetName(String value, String fieldName) { // LIM-003
      ExcelSheetNames.requireValid(value, fieldName);
    }

    static void requireNonNegative(int value, String fieldName) {
      if (value < 0) {
        throw new IllegalArgumentException(fieldName + " must not be negative");
      }
    }

    static void requirePositive(int value, String fieldName) {
      if (value <= 0) {
        throw new IllegalArgumentException(fieldName + " must be greater than 0");
      }
    }

    static void requireNonZero(int value, String fieldName) {
      if (value == 0) {
        throw new IllegalArgumentException(fieldName + " must not be 0");
      }
    }

    static void requireRowIndex(int value, String fieldName) {
      // LIM-008
      requireNonNegative(value, fieldName);
      if (value > ExcelRowSpan.MAX_ROW_INDEX) {
        throw new IllegalArgumentException(
            fieldName + " must not exceed " + ExcelRowSpan.MAX_ROW_INDEX + " (Excel row limit)");
      }
    }

    static void requireColumnIndex(int value, String fieldName) {
      // LIM-009
      requireNonNegative(value, fieldName);
      if (value > ExcelColumnSpan.MAX_COLUMN_INDEX) {
        throw new IllegalArgumentException(
            fieldName
                + " must not exceed "
                + ExcelColumnSpan.MAX_COLUMN_INDEX
                + " (Excel column limit)");
      }
    }

    static void requireOrderedSpan(
        int firstValue, int lastValue, String firstFieldName, String lastFieldName) {
      if (lastValue < firstValue) {
        throw new IllegalArgumentException(
            lastFieldName + " must not be less than " + firstFieldName);
      }
    }

    static void requireColumnWidthCharacters(double widthCharacters) { // LIM-004
      ExcelSheetLayoutLimits.requireColumnWidthCharacters(widthCharacters, "widthCharacters");
    }

    static void requireRowHeightPoints(double heightPoints) { // LIM-005
      ExcelSheetLayoutLimits.requireRowHeightPoints(heightPoints, "heightPoints");
    }

    static void requireNamedRangeName(String name) {
      ProtocolDefinedNameValidation.validateName(name);
    }

    static void requirePivotTableName(String name) {
      ExcelPivotTableNaming.validateName(name);
    }

    static void requireZoomPercent(int zoomPercent) { // LIM-022
      ExcelSheetLayoutLimits.requireZoomPercent(zoomPercent, "zoomPercent");
    }

    static List<List<CellInput>> copyRows(List<List<CellInput>> rows) {
      Objects.requireNonNull(rows, "rows must not be null");
      List<List<CellInput>> copy = new ArrayList<>(rows.size());
      for (List<CellInput> row : rows) {
        copy.add(new ArrayList<>(Objects.requireNonNull(row, "rows must not contain null rows")));
      }
      return java.util.Collections.unmodifiableList(copy);
    }

    static List<List<CellInput>> freezeRows(List<List<CellInput>> rows) {
      return rows.stream().map(List::copyOf).toList();
    }

    static List<String> copySheetNames(List<String> sheetNames, String fieldName) {
      Objects.requireNonNull(sheetNames, fieldName + " must not be null");
      List<String> copy = new ArrayList<>(sheetNames);
      for (String sheetName : copy) {
        requireSheetName(sheetName, fieldName);
      }
      return List.copyOf(copy);
    }

    static void requireDistinct(List<String> values, String fieldName) {
      if (new java.util.LinkedHashSet<>(values).size() != values.size()) {
        throw new IllegalArgumentException(fieldName + " must not contain duplicates");
      }
    }

    static void requireRectangularRows(List<List<CellInput>> rows) {
      if (rows.isEmpty()) {
        throw new IllegalArgumentException("rows must not be empty");
      }
      int expectedWidth = -1;
      for (List<CellInput> row : rows) {
        Objects.requireNonNull(row, "rows must not contain null rows");
        if (row.isEmpty()) {
          throw new IllegalArgumentException("rows must not contain empty rows");
        }
        if (expectedWidth < 0) {
          expectedWidth = row.size();
        } else if (row.size() != expectedWidth) {
          throw new IllegalArgumentException("rows must describe a rectangular matrix");
        }
        for (CellInput value : row) {
          Objects.requireNonNull(value, "rows must not contain null cell values");
        }
      }
    }
  }
}
