package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.ComparisonOperator;

/** Canonical comparison operators reused by validation and conditional-format rule families. */
public enum ExcelComparisonOperator {
  BETWEEN,
  NOT_BETWEEN,
  EQUAL,
  NOT_EQUAL,
  GREATER_THAN,
  LESS_THAN,
  GREATER_OR_EQUAL,
  LESS_OR_EQUAL;

  /** Returns the matching Apache POI comparison-operator constant. */
  public byte toPoiComparisonOperator() {
    return switch (this) {
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

  /** Converts one Apache POI comparison-operator constant into the GridGrind enum. */
  public static ExcelComparisonOperator fromPoiComparisonOperator(int comparisonOperator) {
    return switch (comparisonOperator) {
      case ComparisonOperator.BETWEEN -> BETWEEN;
      case ComparisonOperator.NOT_BETWEEN -> NOT_BETWEEN;
      case ComparisonOperator.EQUAL -> EQUAL;
      case ComparisonOperator.NOT_EQUAL -> NOT_EQUAL;
      case ComparisonOperator.GT -> GREATER_THAN;
      case ComparisonOperator.LT -> LESS_THAN;
      case ComparisonOperator.GE -> GREATER_OR_EQUAL;
      case ComparisonOperator.LE -> LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported Apache POI comparison operator: " + comparisonOperator);
    };
  }
}
