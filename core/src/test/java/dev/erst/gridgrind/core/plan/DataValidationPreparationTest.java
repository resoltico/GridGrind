package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import org.junit.jupiter.api.Test;

/** Tests for the top-level data-validation transport model. */
class DataValidationPreparationTest {
  @Test
  void preparesDefinitionAndNestedPromptFields() {
    DataValidationSpec input =
        dataValidation(
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty()),
            false,
            false,
            prompt("Ship date", "Use today or later.", true),
            errorAlert(
                ExcelDataValidationErrorStyle.WARNING,
                "Invalid date",
                "Date must be today or later.",
                true));

    assertEquals(
        new ResolvedDataValidation(
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", java.util.Optional.empty()),
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
        MutationPreparation.toResolvedDataValidation(input));
  }

  @Test
  void validatesDefinitionAndNestedPromptInputs() {
    assertThrows(
        NullPointerException.class,
        () ->
            new DataValidationSpec(
                null, false, false, java.util.Optional.empty(), java.util.Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationPromptSpec(null, TextSourceSpec.inline("Text"), true));
    assertThrows(IllegalArgumentException.class, () -> prompt("Title", " ", true));
    assertThrows(
        IllegalArgumentException.class,
        () -> errorAlert(ExcelDataValidationErrorStyle.STOP, " ", "Text", true));
    assertThrows(
        NullPointerException.class,
        () ->
            new DataValidationErrorAlertSpec(
                null, TextSourceSpec.inline("Title"), TextSourceSpec.inline("Text"), true));
  }

  @Test
  void nestedPromptAndErrorInputsPreserveVisibilityFlags() {
    assertEquals(
        new ResolvedValidationPrompt("Ship date", "Use today or later.", true),
        MutationPreparation.toResolvedValidationPrompt(
                prompt("Ship date", "Use today or later.", true))
            .orElseThrow());
    assertEquals(
        new ResolvedValidationPrompt("Ship date", "Use today or later.", false),
        MutationPreparation.toResolvedValidationPrompt(
                prompt("Ship date", "Use today or later.", false))
            .orElseThrow());
    assertEquals(
        new ResolvedValidationErrorAlert(
            ExcelDataValidationErrorStyle.INFORMATION, "Invalid date", "Pick a real date.", true),
        MutationPreparation.toResolvedValidationErrorAlert(
                errorAlert(
                    ExcelDataValidationErrorStyle.INFORMATION,
                    "Invalid date",
                    "Pick a real date.",
                    true))
            .orElseThrow());
    assertEquals(
        new ResolvedValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false),
        MutationPreparation.toResolvedValidationErrorAlert(
                errorAlert(
                    ExcelDataValidationErrorStyle.STOP, "Invalid date", "Pick a real date.", false))
            .orElseThrow());
  }

  private static DataValidationPromptSpec prompt(String title, String text, boolean visible) {
    return new DataValidationPromptSpec(
        TextSourceSpec.inline(title), TextSourceSpec.inline(text), visible);
  }

  private static DataValidationErrorAlertSpec errorAlert(
      ExcelDataValidationErrorStyle style, String title, String text, boolean visible) {
    return new DataValidationErrorAlertSpec(
        style, TextSourceSpec.inline(title), TextSourceSpec.inline(text), visible);
  }

  private static DataValidationSpec dataValidation(
      DataValidationRuleSpec rule,
      boolean allowBlank,
      boolean suppressDropDownArrow,
      DataValidationPromptSpec prompt,
      DataValidationErrorAlertSpec errorAlert) {
    return new DataValidationSpec(
        rule,
        allowBlank,
        suppressDropDownArrow,
        java.util.Optional.of(prompt),
        java.util.Optional.of(errorAlert));
  }
}
