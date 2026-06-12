package dev.erst.gridgrind.contract.selector;

import dev.erst.gridgrind.contract.dto.ProtocolCellAddressValidation;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelPivotTableNaming;
import dev.erst.gridgrind.excel.foundation.ExcelReadLimits;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;

/** Shared scalar validation for selector families. */
final class SelectorValueValidation {
  private SelectorValueValidation() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static String requireSheetName(String value, String fieldName) {
    String validated = requireNonBlank(value, fieldName);
    ExcelSheetNames.requireValid(validated, fieldName);
    return validated;
  }

  static String requireDefinedName(String value, String fieldName) {
    try {
      return ProtocolDefinedNameValidation.validateName(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(
          prefixedValidationMessage(fieldName, Objects.toString(exception.getMessage(), "")),
          exception);
    }
  }

  static String requirePivotTableName(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    try {
      return ExcelPivotTableNaming.validateName(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(
          prefixedValidationMessage(fieldName, Objects.toString(exception.getMessage(), "")),
          exception);
    }
  }

  static String requireAddress(String value, String fieldName) {
    try {
      return ProtocolCellAddressValidation.validateAddress(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(
          prefixedValidationMessage(fieldName, Objects.toString(exception.getMessage(), "")),
          exception);
    }
  }

  static String requireRange(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    String[] parts = value.split(":", -1);
    if (parts.length > 2) {
      throw new IllegalArgumentException(
          fieldName + " must be a rectangular A1-style range with at most one ':'");
    }
    requireAddress(parts[0], fieldName);
    if (parts.length == 2) {
      requireAddress(parts[1], fieldName);
    }
    return value;
  }

  static int requirePositive(int value, String fieldName) {
    if (value <= 0) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
    return value;
  }

  static int requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
    return value;
  }

  static int requireNonZero(int value, String fieldName) {
    if (value == 0) {
      throw new IllegalArgumentException(fieldName + " must not be 0");
    }
    return value;
  }

  static int requireRowIndexWithinBounds(int value, String fieldName) {
    // LIM-008
    requireNonNegative(value, fieldName);
    if (value > ExcelRowSpan.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(fieldName + " must be within Excel .xlsx row bounds");
    }
    return value;
  }

  static int requireColumnIndexWithinBounds(int value, String fieldName) {
    // LIM-009
    requireNonNegative(value, fieldName);
    if (value > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(fieldName + " must be within Excel .xlsx column bounds");
    }
    return value;
  }

  static void requireWindowSize(int rowCount, int columnCount) { // LIM-001
    long cells = (long) rowCount * columnCount;
    if (cells > ExcelReadLimits.MAX_WINDOW_CELLS) {
      throw new IllegalArgumentException(
          "rowCount * columnCount must not exceed "
              + ExcelReadLimits.MAX_WINDOW_CELLS
              + " but was "
              + cells);
    }
  }

  static String prefixedValidationMessage(String fieldName, String message) {
    if (message == null || message.isBlank() || message.startsWith(fieldName + " ")) {
      return message;
    }
    return fieldName + " " + message;
  }
}
