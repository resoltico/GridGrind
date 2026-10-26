package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for data-validation rule transport models and engine conversion. */
class DataValidationRuleSpecTest {
  @Test
  void explicitListCopiesValuesAndConvertsToEngineRule() {
    List<String> values = new ArrayList<>(List.of("Queued", "Done"));

    DataValidationRuleSpec.ExplicitList explicitList =
        new DataValidationRuleSpec.ExplicitList(values);
    values.clear();

    assertEquals(List.of("Queued", "Done"), explicitList.values());
    assertEquals(
        new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(explicitList));
  }

  @Test
  void explicitListAllowsEmptyValuesForParityPreservation() {
    DataValidationRuleSpec.ExplicitList explicitList =
        new DataValidationRuleSpec.ExplicitList(List.of());

    assertEquals(List.of(), explicitList.values());
    assertEquals(
        new ExcelDataValidationRule.ExplicitList(List.of()),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(explicitList));
  }

  @Test
  void comparisonRulesDelegateConditionalValidationToEngineRuleConstruction() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new DataValidationRuleSpec.WholeNumber(
                    ExcelComparisonOperator.BETWEEN, "1", java.util.Optional.empty()));

    assertTrue(failure.getMessage().contains("formula2"));
  }

  @Test
  void buildsSupportedRuleFamilies() {
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.GREATER_THAN, "1", java.util.Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.FormulaList("Statuses")));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", java.util.Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "0.5", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", java.util.Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.TimeRule(
                ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL, "20", java.util.Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.TextLength(
                ExcelComparisonOperator.LESS_OR_EQUAL, "20", java.util.Optional.empty())));
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
        WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(
            new DataValidationRuleSpec.CustomFormula("LEN(A1)>0")));
  }

  @Test
  void validatesRuleInputs() {
    assertThrows(NullPointerException.class, () -> new DataValidationRuleSpec.ExplicitList(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.ExplicitList(List.of("Queued", " ")));
    assertThrows(IllegalArgumentException.class, () -> new DataValidationRuleSpec.FormulaList(" "));
    assertThrows(
        NullPointerException.class, () -> new DataValidationRuleSpec.WholeNumber(null, "1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.DecimalNumber(ExcelComparisonOperator.LESS_THAN, " ", null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationRuleSpec.DateRule(null, "TODAY()", null));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationRuleSpec.TimeRule(null, "TIME(9,0,0)", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.TextLength(ExcelComparisonOperator.EQUAL, " ", null));
    assertThrows(NullPointerException.class, () -> new DataValidationRuleSpec.CustomFormula(null));
  }
}
