package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptReport;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;

/** Expected observations for fixtures authored through POI validation APIs. */
public final class DataValidationFixtures {
  private DataValidationFixtures() {}

  /** Selects factual expected fields without sharing the production read path. */
  public static DataValidationEntryReport.DataValidationDefinitionReport report(
      ResolvedDataValidation value) {
    return new DataValidationEntryReport.DataValidationDefinitionReport(
        value.rule(),
        value.allowBlank(),
        value.suppressDropDownArrow(),
        value
            .prompt()
            .map(
                prompt ->
                    new DataValidationPromptReport(
                        prompt.title(), prompt.text(), prompt.showPromptBox())),
        value
            .errorAlert()
            .map(
                error ->
                    new DataValidationErrorAlertReport(
                        error.style(), error.title(), error.text(), error.showErrorBox())));
  }
}
