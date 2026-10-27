package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.problem.FieldValidationBasicRule;
import dev.erst.gridgrind.core.problem.FieldValidationBoundRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.Optional;

/** Numeric and size validation for selector families. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_CELL_RETURNING_READ_COUNT,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_ROWS_PER_WORKSHEET,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_COLUMNS_PER_WORKSHEET
})
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
    if (value > Limit.MAX_ROW_INDEX) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.ROW_INDEX_BOUNDS,
              Integer.toString(Limit.MAX_ROW_INDEX)));
    }
    return value;
  }

  static int requireColumnIndexWithinBounds(int value, String fieldName) {
    requireNonNegative(value, fieldName);
    if (value > Limit.MAX_COLUMN_INDEX) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.COLUMN_INDEX_BOUNDS,
              Integer.toString(Limit.MAX_COLUMN_INDEX)));
    }
    return value;
  }

  static void requireWindowSize(int rowCount, int columnCount) {
    long cells = (long) rowCount * columnCount;
    if (cells > Limit.MAX_READ_CELLS) {
      throw invalidWindowSize(cells);
    }
  }

  static void requireReadCellCount(int count, String fieldName) {
    if (count > Limit.MAX_READ_CELLS) {
      throw SelectorTextValidation.invalidField(
          FieldValidationProblem.atField(
              fieldName,
              FieldValidationBoundRule.COLLECTION_SIZE_LIMIT,
              Integer.toString(Limit.MAX_READ_CELLS),
              Integer.toString(count)));
    }
  }

  private static InvalidRequestException invalidWindowSize(long actualCells) {
    return new InvalidRequestException(
        FieldValidationProblem.detached(
            "rowCount * columnCount",
            FieldValidationBoundRule.WINDOW_SIZE_PRODUCT,
            Long.toString(Limit.MAX_READ_CELLS),
            Long.toString(actualCells)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        null);
  }
}
