package dev.erst.gridgrind.excel.validation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.junit.jupiter.api.Test;

/** Tests the data-validation-specific Apache POI comparison-operator bridge. */
class ExcelDataValidationComparisonOperatorPoiBridgeTest {
  @Test
  void roundTripsEveryDataValidationOperator() {
    for (ExcelComparisonOperator operator : ExcelComparisonOperator.values()) {
      assertEquals(
          operator,
          ExcelDataValidationComparisonOperatorPoiBridge.fromPoi(
              ExcelDataValidationComparisonOperatorPoiBridge.toPoi(operator)));
    }
  }

  @Test
  void readsEveryOoxmlOperatorWithoutInversion() {
    assertEquals(
        ExcelComparisonOperator.BETWEEN,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("between"));
    assertEquals(
        ExcelComparisonOperator.NOT_BETWEEN,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("notBetween"));
    assertEquals(
        ExcelComparisonOperator.EQUAL,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("equal"));
    assertEquals(
        ExcelComparisonOperator.NOT_EQUAL,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("notEqual"));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("greaterThan"));
    assertEquals(
        ExcelComparisonOperator.LESS_THAN,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("lessThan"));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("greaterThanOrEqual"));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL,
        ExcelDataValidationComparisonOperatorPoiBridge.fromXml("lessThanOrEqual"));
  }

  @Test
  void rejectsUnknownOperatorValues() {
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelDataValidationComparisonOperatorPoiBridge.fromPoi(-1));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelDataValidationComparisonOperatorPoiBridge.fromXml("unsupported"));
    assertEquals(
        DataValidationConstraint.OperatorType.LESS_OR_EQUAL,
        ExcelDataValidationComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.LESS_OR_EQUAL));
  }
}
