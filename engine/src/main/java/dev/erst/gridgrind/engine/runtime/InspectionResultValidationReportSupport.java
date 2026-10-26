package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.util.Optional;

/** Converts validation and conditional-format workbook snapshots into protocol reports. */
final class InspectionResultValidationReportSupport {
  private InspectionResultValidationReportSupport() {}

  static DataValidationEntryReport toDataValidationEntryReport(
      ExcelDataValidationSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelDataValidationSnapshot.Supported supported ->
          new DataValidationEntryReport.Supported(
              supported.ranges().stream().map(RangeLabel::new).toList(),
              toDataValidationDefinitionReport(supported.validation()));
      case ExcelDataValidationSnapshot.Unsupported unsupported ->
          new DataValidationEntryReport.Unsupported(
              unsupported.ranges().stream().map(RangeLabel::new).toList(),
              unsupported.kind(),
              unsupported.detail());
    };
  }

  static DataValidationEntryReport.DataValidationDefinitionReport toDataValidationDefinitionReport(
      ResolvedDataValidation definition) {
    return new DataValidationEntryReport.DataValidationDefinitionReport(
        definition.rule(),
        definition.allowBlank(),
        definition.suppressDropDownArrow(),
        promptInput(definition),
        errorAlertInput(definition));
  }

  private static Optional<DataValidationPromptSpec> promptInput(ResolvedDataValidation definition) {
    return definition
        .prompt()
        .map(
            prompt ->
                new DataValidationPromptSpec(
                    new TextSourceSpec.Inline(prompt.title()),
                    new TextSourceSpec.Inline(prompt.text()),
                    prompt.showPromptBox()));
  }

  private static Optional<DataValidationErrorAlertSpec> errorAlertInput(
      ResolvedDataValidation definition) {
    return definition
        .errorAlert()
        .map(
            errorAlert ->
                new DataValidationErrorAlertSpec(
                    errorAlert.style(),
                    new TextSourceSpec.Inline(errorAlert.title()),
                    new TextSourceSpec.Inline(errorAlert.text()),
                    errorAlert.showErrorBox()));
  }
}
