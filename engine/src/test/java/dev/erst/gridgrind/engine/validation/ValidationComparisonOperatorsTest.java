package dev.erst.gridgrind.engine.validation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.junit.jupiter.api.Test;

/** Tests the data-validation-specific Apache POI comparison-operator bridge. */
class ValidationComparisonOperatorsTest {
  @Test
  void roundTripsEveryDataValidationOperator() {
    for (ExcelComparisonOperator operator : ExcelComparisonOperator.values()) {
      assertEquals(operator, DataValidationAdapter.fromPoi(DataValidationAdapter.toPoi(operator)));
    }
  }

  @Test
  void readsEveryOoxmlOperatorWithoutInversion() {
    assertEquals(ExcelComparisonOperator.BETWEEN, DataValidationAdapter.fromXml("between"));
    assertEquals(ExcelComparisonOperator.NOT_BETWEEN, DataValidationAdapter.fromXml("notBetween"));
    assertEquals(ExcelComparisonOperator.EQUAL, DataValidationAdapter.fromXml("equal"));
    assertEquals(ExcelComparisonOperator.NOT_EQUAL, DataValidationAdapter.fromXml("notEqual"));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN, DataValidationAdapter.fromXml("greaterThan"));
    assertEquals(ExcelComparisonOperator.LESS_THAN, DataValidationAdapter.fromXml("lessThan"));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        DataValidationAdapter.fromXml("greaterThanOrEqual"));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL, DataValidationAdapter.fromXml("lessThanOrEqual"));
  }

  @Test
  void rejectsUnknownOperatorValues() {
    assertThrows(IllegalArgumentException.class, () -> DataValidationAdapter.fromPoi(-1));
    assertThrows(
        IllegalArgumentException.class, () -> DataValidationAdapter.fromXml("unsupported"));
    assertEquals(
        DataValidationConstraint.OperatorType.LESS_OR_EQUAL,
        DataValidationAdapter.toPoi(ExcelComparisonOperator.LESS_OR_EQUAL));
  }
}
