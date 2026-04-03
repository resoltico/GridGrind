package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.protocol.dto.ComparisonOperator;
import dev.erst.gridgrind.protocol.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.protocol.dto.DataValidationErrorStyle;
import dev.erst.gridgrind.protocol.dto.DataValidationInput;
import dev.erst.gridgrind.protocol.dto.DataValidationPromptInput;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import org.junit.jupiter.api.Test;

/** Tests for the top-level data-validation transport model. */
class DataValidationInputTest {
  @Test
  void defaultsBooleansAndConvertsToEngineDefinition() {
    DataValidationInput input =
        new DataValidationInput(
            new DataValidationRuleInput.DateRule(
                ComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
            null,
            null,
            new DataValidationPromptInput("Ship date", "Use today or later.", null),
            new DataValidationErrorAlertInput(
                DataValidationErrorStyle.WARNING,
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
        DefaultGridGrindRequestExecutor.toExcelDataValidationDefinition(input));
  }

  @Test
  void validatesDefinitionAndNestedPromptInputs() {
    assertThrows(
        NullPointerException.class, () -> new DataValidationInput(null, false, false, null, null));
    assertThrows(
        NullPointerException.class, () -> new DataValidationPromptInput(null, "Text", true));
    assertThrows(
        IllegalArgumentException.class, () -> new DataValidationPromptInput("Title", " ", true));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationErrorAlertInput(DataValidationErrorStyle.STOP, " ", "Text", true));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationErrorAlertInput(null, "Title", "Text", true));
  }

  @Test
  void nestedPromptAndErrorInputsDefaultAndPreserveVisibilityFlags() {
    assertEquals(
        new ExcelDataValidationPrompt("Ship date", "Use today or later.", true),
        DefaultGridGrindRequestExecutor.toExcelDataValidationPrompt(
            new DataValidationPromptInput("Ship date", "Use today or later.", null)));
    assertEquals(
        new ExcelDataValidationPrompt("Ship date", "Use today or later.", false),
        DefaultGridGrindRequestExecutor.toExcelDataValidationPrompt(
            new DataValidationPromptInput("Ship date", "Use today or later.", false)));
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.INFORMATION, "Invalid date", "Pick a real date.", true),
        DefaultGridGrindRequestExecutor.toExcelDataValidationErrorAlert(
            new DataValidationErrorAlertInput(
                DataValidationErrorStyle.INFORMATION, "Invalid date", "Pick a real date.", null)));
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false),
        DefaultGridGrindRequestExecutor.toExcelDataValidationErrorAlert(
            new DataValidationErrorAlertInput(
                DataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false)));
  }
}
