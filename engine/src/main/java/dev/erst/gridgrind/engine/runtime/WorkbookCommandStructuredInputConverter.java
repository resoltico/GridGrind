package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import java.util.List;
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

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      List<String> ranges, ConditionalFormattingSpec definition) {
    return new ExcelConditionalFormattingBlockDefinition(ranges, definition.rules());
  }

  static ExcelAutofilterFilterColumn toExcelAutofilterFilterColumn(
      AutofilterFilterColumnSpec column) {
    return new ExcelAutofilterFilterColumn(
        column.columnId(),
        column.showButton(),
        toExcelAutofilterFilterCriterion(column.criterion()));
  }

  static Optional<ExcelAutofilterSortState> toExcelAutofilterSortState(
      AutofilterSortStateSpec sortState) {
    if (sortState == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelAutofilterSortState(
            sortState.range().value(),
            sortState.caseSensitive(),
            sortState.columnSort(),
            sortState.sortMethod(),
            sortState.conditions().stream()
                .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterSortCondition)
                .toList()));
  }

  private static ExcelAutofilterFilterCriterion toExcelAutofilterFilterCriterion(
      AutofilterFilterCriterionSpec criterion) {
    return switch (criterion) {
      case AutofilterFilterCriterionSpec.Values values ->
          new ExcelAutofilterFilterCriterion.Values(values.values(), values.includeBlank());
      case AutofilterFilterCriterionSpec.Custom custom ->
          new ExcelAutofilterFilterCriterion.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new ExcelAutofilterFilterCriterion.CustomCondition(
                              condition.operator(), condition.value()))
                  .toList());
      case AutofilterFilterCriterionSpec.Dynamic dynamic ->
          new ExcelAutofilterFilterCriterion.Dynamic(
              dynamic.dynamicType(), dynamic.value().orElse(null), dynamic.maxValue().orElse(null));
      case AutofilterFilterCriterionSpec.Top10 top10 ->
          new ExcelAutofilterFilterCriterion.Top10(top10.value(), top10.top(), top10.percent());
      case AutofilterFilterCriterionSpec.Color color ->
          new ExcelAutofilterFilterCriterion.Color(color.cellColor(), color.color());
      case AutofilterFilterCriterionSpec.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static ExcelAutofilterSortCondition toExcelAutofilterSortCondition(
      AutofilterSortConditionSpec condition) {
    return switch (condition) {
      case AutofilterSortConditionSpec.Value value ->
          new ExcelAutofilterSortCondition.Value(value.range().value(), value.descending());
      case AutofilterSortConditionSpec.CellColor cellColor ->
          new ExcelAutofilterSortCondition.CellColor(
              cellColor.range().value(), cellColor.descending(), cellColor.color());
      case AutofilterSortConditionSpec.FontColor fontColor ->
          new ExcelAutofilterSortCondition.FontColor(
              fontColor.range().value(), fontColor.descending(), fontColor.color());
      case AutofilterSortConditionSpec.Icon icon ->
          new ExcelAutofilterSortCondition.Icon(
              icon.range().value(), icon.descending(), icon.iconId());
    };
  }
}
