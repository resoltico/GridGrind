package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import java.util.Optional;

/** Converts structured contract inputs such as drawings, validations, and tables. */
final class WorkbookCommandStructuredInputConverter {
  private WorkbookCommandStructuredInputConverter() {}

  static ResolvedDataValidation toResolvedDataValidation(DataValidationSpec validation) {
    return new ResolvedDataValidation(
        validation.rule(),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        validation
            .prompt()
            .flatMap(WorkbookCommandStructuredInputConverter::toResolvedValidationPrompt),
        validation
            .errorAlert()
            .flatMap(WorkbookCommandStructuredInputConverter::toResolvedValidationErrorAlert));
  }

  static Optional<ResolvedValidationPrompt> toResolvedValidationPrompt(
      DataValidationPromptSpec prompt) {
    return prompt == null
        ? Optional.empty()
        : Optional.of(
            new ResolvedValidationPrompt(
                WorkbookCommandSourceSupport.inlineText(prompt.title(), "validation prompt title"),
                WorkbookCommandSourceSupport.inlineText(prompt.text(), "validation prompt text"),
                prompt.showPromptBox()));
  }

  static Optional<ResolvedValidationErrorAlert> toResolvedValidationErrorAlert(
      DataValidationErrorAlertSpec errorAlert) {
    return errorAlert == null
        ? Optional.empty()
        : Optional.of(
            new ResolvedValidationErrorAlert(
                errorAlert.style(),
                WorkbookCommandSourceSupport.inlineText(
                    errorAlert.title(), "validation error title"),
                WorkbookCommandSourceSupport.inlineText(errorAlert.text(), "validation error text"),
                errorAlert.showErrorBox()));
  }
}
