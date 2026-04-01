package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for data-validation rule transport models and engine conversion. */
class DataValidationRuleInputTest {
  @Test
  void explicitListCopiesValuesAndConvertsToEngineRule() {
    List<String> values = new ArrayList<>(List.of("Queued", "Done"));

    DataValidationRuleInput.ExplicitList explicitList =
        new DataValidationRuleInput.ExplicitList(values);
    values.clear();

    assertEquals(List.of("Queued", "Done"), explicitList.values());
    assertEquals(
        new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
        explicitList.toExcelDataValidationRule());
  }

  @Test
  void comparisonRulesDelegateConditionalValidationToEngineRuleConstruction() {
    DataValidationRuleInput.WholeNumber between =
        new DataValidationRuleInput.WholeNumber(ExcelComparisonOperator.BETWEEN, "1", null);

    IllegalArgumentException failure =
        assertThrows(IllegalArgumentException.class, between::toExcelDataValidationRule);

    assertTrue(failure.getMessage().contains("formula2"));
  }

  @Test
  void buildsSupportedRuleFamilies() {
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.GREATER_THAN, "1", null),
        new DataValidationRuleInput.WholeNumber(ExcelComparisonOperator.GREATER_THAN, "1", null)
            .toExcelDataValidationRule());
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        new DataValidationRuleInput.FormulaList("Statuses").toExcelDataValidationRule());
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", null),
        new DataValidationRuleInput.DecimalNumber(ExcelComparisonOperator.GREATER_THAN, "0.5", null)
            .toExcelDataValidationRule());
    assertEquals(
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
        new DataValidationRuleInput.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null)
            .toExcelDataValidationRule());
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null),
        new DataValidationRuleInput.TimeRule(
                ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null)
            .toExcelDataValidationRule());
    assertEquals(
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        new DataValidationRuleInput.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null)
            .toExcelDataValidationRule());
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
        new DataValidationRuleInput.CustomFormula("LEN(A1)>0").toExcelDataValidationRule());
  }

  @Test
  void validatesRuleInputs() {
    assertThrows(NullPointerException.class, () -> new DataValidationRuleInput.ExplicitList(null));
    assertThrows(
        IllegalArgumentException.class, () -> new DataValidationRuleInput.ExplicitList(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.ExplicitList(List.of("Queued", " ")));
    assertThrows(
        IllegalArgumentException.class, () -> new DataValidationRuleInput.FormulaList(" "));
    assertThrows(
        NullPointerException.class, () -> new DataValidationRuleInput.WholeNumber(null, "1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.DecimalNumber(
                ExcelComparisonOperator.LESS_THAN, " ", null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationRuleInput.DateRule(null, "TODAY()", null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationRuleInput.TimeRule(null, "TIME(9,0,0)", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.TextLength(ExcelComparisonOperator.EQUAL, " ", null));
    assertThrows(NullPointerException.class, () -> new DataValidationRuleInput.CustomFormula(null));
  }
}
