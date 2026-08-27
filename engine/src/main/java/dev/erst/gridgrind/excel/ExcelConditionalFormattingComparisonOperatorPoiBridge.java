package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import org.apache.poi.ss.usermodel.ComparisonOperator;

/** Maps conditional-formatting comparison operators to Apache POI constants. */
public final class ExcelConditionalFormattingComparisonOperatorPoiBridge {
  private ExcelConditionalFormattingComparisonOperatorPoiBridge() {}

  /** Returns the POI conditional-formatting operator constant for one protocol operator. */
  public static byte toPoi(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> ComparisonOperator.BETWEEN;
      case NOT_BETWEEN -> ComparisonOperator.NOT_BETWEEN;
      case EQUAL -> ComparisonOperator.EQUAL;
      case NOT_EQUAL -> ComparisonOperator.NOT_EQUAL;
      case GREATER_THAN -> ComparisonOperator.GT;
      case LESS_THAN -> ComparisonOperator.LT;
      case GREATER_OR_EQUAL -> ComparisonOperator.GE;
      case LESS_OR_EQUAL -> ComparisonOperator.LE;
    };
  }

  /** Returns the protocol operator represented by one POI conditional-formatting constant. */
  public static ExcelComparisonOperator fromPoi(int comparisonOperator) {
    return switch (comparisonOperator) {
      case ComparisonOperator.BETWEEN -> ExcelComparisonOperator.BETWEEN;
      case ComparisonOperator.NOT_BETWEEN -> ExcelComparisonOperator.NOT_BETWEEN;
      case ComparisonOperator.EQUAL -> ExcelComparisonOperator.EQUAL;
      case ComparisonOperator.NOT_EQUAL -> ExcelComparisonOperator.NOT_EQUAL;
      case ComparisonOperator.GT -> ExcelComparisonOperator.GREATER_THAN;
      case ComparisonOperator.LT -> ExcelComparisonOperator.LESS_THAN;
      case ComparisonOperator.GE -> ExcelComparisonOperator.GREATER_OR_EQUAL;
      case ComparisonOperator.LE -> ExcelComparisonOperator.LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported Apache POI conditional-formatting comparison operator: "
                  + comparisonOperator);
    };
  }
}
