package dev.erst.gridgrind.excel.validation;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidationConstraint;

/** Maps data-validation comparison operators to Apache POI constants. */
final class ExcelDataValidationComparisonOperatorPoiBridge {
  private ExcelDataValidationComparisonOperatorPoiBridge() {}

  static int toPoi(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> DataValidationConstraint.OperatorType.BETWEEN;
      case NOT_BETWEEN -> DataValidationConstraint.OperatorType.NOT_BETWEEN;
      case EQUAL -> DataValidationConstraint.OperatorType.EQUAL;
      case NOT_EQUAL -> DataValidationConstraint.OperatorType.NOT_EQUAL;
      case GREATER_THAN -> DataValidationConstraint.OperatorType.GREATER_THAN;
      case LESS_THAN -> DataValidationConstraint.OperatorType.LESS_THAN;
      case GREATER_OR_EQUAL -> DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
      case LESS_OR_EQUAL -> DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
    };
  }

  static ExcelComparisonOperator fromPoi(int operator) {
    return switch (operator) {
      case DataValidationConstraint.OperatorType.BETWEEN -> ExcelComparisonOperator.BETWEEN;
      case DataValidationConstraint.OperatorType.NOT_BETWEEN -> ExcelComparisonOperator.NOT_BETWEEN;
      case DataValidationConstraint.OperatorType.EQUAL -> ExcelComparisonOperator.EQUAL;
      case DataValidationConstraint.OperatorType.NOT_EQUAL -> ExcelComparisonOperator.NOT_EQUAL;
      case DataValidationConstraint.OperatorType.GREATER_THAN ->
          ExcelComparisonOperator.GREATER_THAN;
      case DataValidationConstraint.OperatorType.LESS_THAN -> ExcelComparisonOperator.LESS_THAN;
      case DataValidationConstraint.OperatorType.GREATER_OR_EQUAL ->
          ExcelComparisonOperator.GREATER_OR_EQUAL;
      case DataValidationConstraint.OperatorType.LESS_OR_EQUAL ->
          ExcelComparisonOperator.LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported Apache POI data-validation comparison operator: " + operator);
    };
  }

  static ExcelComparisonOperator fromXml(String operator) {
    return switch (operator) {
      case "between" -> ExcelComparisonOperator.BETWEEN;
      case "notBetween" -> ExcelComparisonOperator.NOT_BETWEEN;
      case "equal" -> ExcelComparisonOperator.EQUAL;
      case "notEqual" -> ExcelComparisonOperator.NOT_EQUAL;
      case "greaterThan" -> ExcelComparisonOperator.GREATER_THAN;
      case "lessThan" -> ExcelComparisonOperator.LESS_THAN;
      case "greaterThanOrEqual" -> ExcelComparisonOperator.GREATER_OR_EQUAL;
      case "lessThanOrEqual" -> ExcelComparisonOperator.LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported OOXML data-validation comparison operator: " + operator);
    };
  }
}
