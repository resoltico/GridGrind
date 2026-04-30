package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import org.junit.jupiter.api.Test;

/** Tests for the top-level data-validation transport model. */
class DataValidationInputTest {
  @Test
  void convertsToEngineDefinitionAndNestedPromptDefaults() {
    DataValidationInput input =
        new DataValidationInput(
            new DataValidationRuleInput.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
            false,
            false,
            prompt("Ship date", "Use today or later.", null),
            errorAlert(
                ExcelDataValidationErrorStyle.WARNING,
                "Invalid date",
                "Date must be today or later.",
                null));

    assertEquals(
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
            false,
            false,
            new ExcelDataValidationPrompt("Ship date", "Use today or later.", true),
            new ExcelDataValidationErrorAlert(
                ExcelDataValidationErrorStyle.WARNING,
                "Invalid date",
                "Date must be today or later.",
                true)),
        WorkbookCommandConverter.toExcelDataValidationDefinition(input));
  }

  @Test
  void validatesDefinitionAndNestedPromptInputs() {
    assertThrows(
        NullPointerException.class, () -> new DataValidationInput(null, false, false, null, null));
    assertThrows(
        NullPointerException.class, () -> new DataValidationPromptInput(null, text("Text"), true));
    assertThrows(IllegalArgumentException.class, () -> prompt("Title", " ", true));
    assertThrows(
        IllegalArgumentException.class,
        () -> errorAlert(ExcelDataValidationErrorStyle.STOP, " ", "Text", true));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationErrorAlertInput(null, text("Title"), text("Text"), true));
  }

  @Test
  void nestedPromptAndErrorInputsDefaultAndPreserveVisibilityFlags() {
    assertEquals(
        new ExcelDataValidationPrompt("Ship date", "Use today or later.", true),
        WorkbookCommandConverter.toExcelDataValidationPrompt(
                prompt("Ship date", "Use today or later.", null))
            .orElseThrow());
    assertEquals(
        new ExcelDataValidationPrompt("Ship date", "Use today or later.", false),
        WorkbookCommandConverter.toExcelDataValidationPrompt(
                prompt("Ship date", "Use today or later.", false))
            .orElseThrow());
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.INFORMATION, "Invalid date", "Pick a real date.", true),
        WorkbookCommandConverter.toExcelDataValidationErrorAlert(
                errorAlert(
                    ExcelDataValidationErrorStyle.INFORMATION,
                    "Invalid date",
                    "Pick a real date.",
                    null))
            .orElseThrow());
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false),
        WorkbookCommandConverter.toExcelDataValidationErrorAlert(
                errorAlert(
                    ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false))
            .orElseThrow());
  }
}
