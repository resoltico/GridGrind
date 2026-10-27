package dev.erst.gridgrind.core.model.validation;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Checks canonical rule bounds, list ownership and explicit-empty-list preservation. */
class DataValidationRuleSpecTest {
  @Test
  void explicitListCopiesValues() {
    List<String> values = new ArrayList<>(List.of("Queued", "Done"));

    DataValidationRuleSpec.ExplicitList explicitList =
        new DataValidationRuleSpec.ExplicitList(values);
    values.clear();

    assertEquals(List.of("Queued", "Done"), explicitList.values());
  }

  @Test
  void explicitListAllowsEmptyValuesForParityPreservation() {
    DataValidationRuleSpec.ExplicitList explicitList =
        new DataValidationRuleSpec.ExplicitList(List.of());

    assertEquals(List.of(), explicitList.values());
  }

  @Test
  void comparisonRulesRequireTheDeclaredUpperBound() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new DataValidationRuleSpec.WholeNumber(
                    ExcelComparisonOperator.BETWEEN, "1", java.util.Optional.empty()));

    assertTrue(failure.getMessage().contains("formula2"));
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
