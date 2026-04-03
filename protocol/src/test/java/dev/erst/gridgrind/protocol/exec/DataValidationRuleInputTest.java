package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.protocol.dto.ComparisonOperator;
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
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(explicitList));
  }

  @Test
  void comparisonRulesDelegateConditionalValidationToEngineRuleConstruction() {
    DataValidationRuleInput.WholeNumber between =
        new DataValidationRuleInput.WholeNumber(ComparisonOperator.BETWEEN, "1", null);

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> DefaultGridGrindRequestExecutor.toExcelDataValidationRule(between));

    assertTrue(failure.getMessage().contains("formula2"));
  }

  @Test
  void buildsSupportedRuleFamilies() {
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.GREATER_THAN, "1", null),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.WholeNumber(ComparisonOperator.GREATER_THAN, "1", null)));
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.FormulaList("Statuses")));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", null),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.DecimalNumber(
                ComparisonOperator.GREATER_THAN, "0.5", null)));
    assertEquals(
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.DateRule(
                ComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null)));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.TimeRule(
                ComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null)));
    assertEquals(
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.TextLength(ComparisonOperator.LESS_OR_EQUAL, "20", null)));
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
        DefaultGridGrindRequestExecutor.toExcelDataValidationRule(
            new DataValidationRuleInput.CustomFormula("LEN(A1)>0")));
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
        () -> new DataValidationRuleInput.DecimalNumber(ComparisonOperator.LESS_THAN, " ", null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationRuleInput.DateRule(null, "TODAY()", null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationRuleInput.TimeRule(null, "TIME(9,0,0)", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.TextLength(ComparisonOperator.EQUAL, " ", null));
    assertThrows(NullPointerException.class, () -> new DataValidationRuleInput.CustomFormula(null));
  }
}
