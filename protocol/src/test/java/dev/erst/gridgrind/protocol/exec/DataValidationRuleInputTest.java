package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
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
    DataValidationRuleInput.WholeNumber between =
        new DataValidationRuleInput.WholeNumber(ExcelComparisonOperator.BETWEEN, "1", null);

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> WorkbookCommandConverter.toExcelDataValidationRule(between));

    assertTrue(failure.getMessage().contains("formula2"));
  }

  @Test
  void buildsSupportedRuleFamilies() {
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.GREATER_THAN, "1", null),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", null)));
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.FormulaList("Statuses")));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", null),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "0.5", null)));
    assertEquals(
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null)));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.TimeRule(
                ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null)));
    assertEquals(
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        WorkbookCommandConverter.toExcelDataValidationRule(
            new DataValidationRuleInput.TextLength(
                ExcelComparisonOperator.LESS_OR_EQUAL, "20", null)));
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
