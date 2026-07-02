package dev.erst.gridgrind.contract.selector;

import dev.erst.gridgrind.contract.json.FieldValidationBasicRule;
import dev.erst.gridgrind.contract.json.FieldValidationBoundRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelReadLimits;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.Optional;

/** Numeric and size validation for selector families. */
final class SelectorNumberValidation {
  private SelectorNumberValidation() {}

  static int requirePositive(int value, String fieldName) {
    if (value <= 0) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.GREATER_THAN_ZERO));
    }
    return value;
  }

  static int requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_NEGATIVE));
    }
    return value;
  }

  static int requireNonZero(int value, String fieldName) {
    if (value == 0) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(fieldName, FieldValidationBasicRule.NON_ZERO));
    }
    return value;
  }

  static int requireRowIndexWithinBounds(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > ExcelRowSpan.MAX_ROW_INDEX) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.ROW_INDEX_BOUNDS,
              Integer.toString(ExcelRowSpan.MAX_ROW_INDEX)));
    }
    return value;
  }

  static int requireColumnIndexWithinBounds(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.COLUMN_INDEX_BOUNDS,
              Integer.toString(ExcelColumnSpan.MAX_COLUMN_INDEX)));
    }
    return value;
  }

  static void requireWindowSize(int rowCount, int columnCount) {
    long cells = (long) rowCount * columnCount;
    if (cells > ExcelReadLimits.MAX_READ_CELLS) {
      throw invalidWindowSize(cells);
    }
  }

  static void requireReadCellCount(int count, String fieldName) {
    if (count > ExcelReadLimits.MAX_READ_CELLS) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.COLLECTION_SIZE_LIMIT,
              Integer.toString(ExcelReadLimits.MAX_READ_CELLS),
              Integer.toString(count)));
    }
  }

  private static InvalidRequestException invalidWindowSize(long actualCells) {
    return new InvalidRequestException(
        FieldValidationProblem.detached(
            "rowCount * columnCount",
            FieldValidationBoundRule.WINDOW_SIZE_PRODUCT,
            Long.toString(ExcelReadLimits.MAX_READ_CELLS),
            Long.toString(actualCells)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        null);
  }
}
