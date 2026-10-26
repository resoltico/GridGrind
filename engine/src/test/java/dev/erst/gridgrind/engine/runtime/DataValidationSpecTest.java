package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelComparisonFormulaSupport;
import org.junit.jupiter.api.Test;

/** Tests for the top-level data-validation transport model. */
class DataValidationSpecTest {
  @Test
  void convertsToEngineDefinitionAndNestedPromptDefaults() {
    DataValidationSpec input =
        dataValidation(
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty()),
            false,
            false,
            prompt("Ship date", "Use today or later.", null),
            errorAlert(
                ExcelDataValidationErrorStyle.WARNING,
                "Invalid date",
                "Date must be today or later.",
                null));

    assertEquals(
        new ResolvedDataValidation(
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL,
                ExcelComparisonFormulaSupport.normalizeFormula("TODAY()", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    java.util.Optional.empty(), "formula2")),
            false,
            false,
            java.util.Optional.of(
                new ResolvedValidationPrompt("Ship date", "Use today or later.", true)),
            java.util.Optional.of(
                new ResolvedValidationErrorAlert(
                    ExcelDataValidationErrorStyle.WARNING,
                    "Invalid date",
                    "Date must be today or later.",
                    true))),
        WorkbookCommandStructuredInputConverter.toResolvedDataValidation(input));
  }

  @Test
  void validatesDefinitionAndNestedPromptInputs() {
    assertThrows(
        NullPointerException.class,
        () ->
            new DataValidationSpec(
                null, false, false, java.util.Optional.empty(), java.util.Optional.empty()));
    assertThrows(
        NullPointerException.class, () -> new DataValidationPromptSpec(null, text("Text"), true));
    assertThrows(IllegalArgumentException.class, () -> prompt("Title", " ", true));
    assertThrows(
        IllegalArgumentException.class,
        () -> errorAlert(ExcelDataValidationErrorStyle.STOP, " ", "Text", true));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationErrorAlertSpec(null, text("Title"), text("Text"), true));
  }

  @Test
  void nestedPromptAndErrorInputsDefaultAndPreserveVisibilityFlags() {
    assertEquals(
        new ResolvedValidationPrompt("Ship date", "Use today or later.", true),
        WorkbookCommandStructuredInputConverter.toResolvedValidationPrompt(
                prompt("Ship date", "Use today or later.", null))
            .orElseThrow());
    assertEquals(
        new ResolvedValidationPrompt("Ship date", "Use today or later.", false),
        WorkbookCommandStructuredInputConverter.toResolvedValidationPrompt(
                prompt("Ship date", "Use today or later.", false))
            .orElseThrow());
    assertEquals(
        new ResolvedValidationErrorAlert(
            ExcelDataValidationErrorStyle.INFORMATION, "Invalid date", "Pick a real date.", true),
        WorkbookCommandStructuredInputConverter.toResolvedValidationErrorAlert(
                errorAlert(
                    ExcelDataValidationErrorStyle.INFORMATION,
                    "Invalid date",
                    "Pick a real date.",
                    null))
            .orElseThrow());
    assertEquals(
        new ResolvedValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false),
        WorkbookCommandStructuredInputConverter.toResolvedValidationErrorAlert(
                errorAlert(
                    ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false))
            .orElseThrow());
  }
}
