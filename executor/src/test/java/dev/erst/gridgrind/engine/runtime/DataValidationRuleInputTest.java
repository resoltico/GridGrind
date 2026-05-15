package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
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
        WorkbookCommandConverter.toExcelDataValidationRule(explicitList));
  }

  @Test
  void explicitListAllowsEmptyValuesForParityPreservation() {
    DataValidationRuleInput.ExplicitList explicitList =
        new DataValidationRuleInput.ExplicitList(List.of());

    assertEquals(List.of(), explicitList.values());
    assertEquals(
        new ExcelDataValidationRule.ExplicitList(List.of()),
        WorkbookCommandConverter.toExcelDataValidationRule(explicitList));
  }

  @Test
  void comparisonRulesDelegateConditionalValidationToEngineRuleConstruction() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new DataValidationRuleInput.WholeNumber(
                    ExcelComparisonOperator.BETWEEN, "1", java.util.Optional.empty()));

    assertTrue(failure.getMessage().contains("formula2"));
  }

  @Test
  void buildsSupportedRuleFamilies() {
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.GREATER_THAN, "1", java.util.Optional.empty()),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.FormulaList("Statuses")));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", java.util.Optional.empty()),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "0.5", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty()),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", java.util.Optional.empty()),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.TimeRule(
                ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL, "20", java.util.Optional.empty()),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.TextLength(
                ExcelComparisonOperator.LESS_OR_EQUAL, "20", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.CustomFormula("LEN(A1)>0")));
  }

  @Test
  void validatesRuleInputs() {
    assertThrows(NullPointerException.class, () -> new DataValidationRuleInput.ExplicitList(null));
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
