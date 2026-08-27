package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingDefinitionInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
import java.util.List;
import java.util.Optional;

/** Converts structured contract inputs such as drawings, validations, and tables. */
final class WorkbookCommandStructuredInputConverter {
  private WorkbookCommandStructuredInputConverter() {}

  static ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return new ExcelDataValidationDefinition(
        toExcelDataValidationRule(validation.rule()),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        validation
            .prompt()
            .flatMap(WorkbookCommandStructuredInputConverter::toExcelDataValidationPrompt),
        validation
            .errorAlert()
            .flatMap(WorkbookCommandStructuredInputConverter::toExcelDataValidationErrorAlert));
  }

  static ExcelDataValidationRule toExcelDataValidationRule(DataValidationRuleInput rule) {
    return switch (rule) {
      case DataValidationRuleInput.ExplicitList explicitList ->
          new ExcelDataValidationRule.ExplicitList(explicitList.values());
      case DataValidationRuleInput.FormulaList formulaList ->
          new ExcelDataValidationRule.FormulaList(formulaList.formula());
      case DataValidationRuleInput.WholeNumber wholeNumber ->
          new ExcelDataValidationRule.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case DataValidationRuleInput.DecimalNumber decimalNumber ->
          new ExcelDataValidationRule.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case DataValidationRuleInput.DateRule dateRule ->
          new ExcelDataValidationRule.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case DataValidationRuleInput.TimeRule timeRule ->
          new ExcelDataValidationRule.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case DataValidationRuleInput.TextLength textLength ->
          new ExcelDataValidationRule.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case DataValidationRuleInput.CustomFormula customFormula ->
          new ExcelDataValidationRule.CustomFormula(customFormula.formula());
    };
  }

  static Optional<ExcelDataValidationPrompt> toExcelDataValidationPrompt(
      DataValidationPromptInput prompt) {
    return prompt == null
        ? Optional.empty()
        : Optional.of(
            new ExcelDataValidationPrompt(
                WorkbookCommandSourceSupport.inlineText(prompt.title(), "validation prompt title"),
                WorkbookCommandSourceSupport.inlineText(prompt.text(), "validation prompt text"),
                prompt.showPromptBox()));
  }

  static Optional<ExcelDataValidationErrorAlert> toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return errorAlert == null
        ? Optional.empty()
        : Optional.of(
            new ExcelDataValidationErrorAlert(
                errorAlert.style(),
                WorkbookCommandSourceSupport.inlineText(
                    errorAlert.title(), "validation error title"),
                WorkbookCommandSourceSupport.inlineText(errorAlert.text(), "validation error text"),
                errorAlert.showErrorBox()));
  }

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      List<String> ranges, ConditionalFormattingDefinitionInput definition) {
    return new ExcelConditionalFormattingBlockDefinition(
        ranges,
        definition.rules().stream()
            .map(WorkbookCommandStructuredInputConverter::toExcelConditionalFormattingRule)
            .toList());
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return switch (rule) {
      case ConditionalFormattingRuleInput.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              formulaRule
                  .style()
                  .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialStyle));
      case ConditionalFormattingRuleInput.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              cellValueRule
                  .style()
                  .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialStyle));
      case ConditionalFormattingRuleInput.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(
                      WorkbookCommandStructuredInputConverter
                          ::toExcelConditionalFormattingThreshold)
                  .toList(),
              colorScaleRule.colors().stream()
                  .map(
                      color ->
                          WorkbookCommandCellInputConverter.toRequiredExcelColor(
                              color, "color-scale color"))
                  .toList(),
              colorScaleRule.stopIfTrue());
      case ConditionalFormattingRuleInput.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              WorkbookCommandCellInputConverter.toRequiredExcelColor(
                  dataBarRule.color(), "data-bar color"),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toExcelConditionalFormattingThreshold(dataBarRule.minThreshold()),
              toExcelConditionalFormattingThreshold(dataBarRule.maxThreshold()),
              dataBarRule.stopIfTrue());
      case ConditionalFormattingRuleInput.IconSetRule iconSetRule ->
          new ExcelConditionalFormattingRule.IconSetRule(
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(
                      WorkbookCommandStructuredInputConverter
                          ::toExcelConditionalFormattingThreshold)
                  .toList(),
              iconSetRule.stopIfTrue());
      case ConditionalFormattingRuleInput.Top10Rule top10Rule ->
          new ExcelConditionalFormattingRule.Top10Rule(
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              top10Rule.stopIfTrue(),
              top10Rule
                  .style()
                  .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialStyle));
    };
  }

  static Optional<ExcelDifferentialStyle> toExcelDifferentialStyle(DifferentialStyleInput style) {
    if (style == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialStyle(
            style.numberFormat(),
            style.bold(),
            style.italic(),
            style.fontHeight().flatMap(WorkbookCommandCellInputConverter::toExcelFontHeight),
            style.fontColor().flatMap(WorkbookCommandCellInputConverter::toExcelColor),
            style.underline(),
            style.strikeout(),
            style.fillColor().flatMap(WorkbookCommandCellInputConverter::toExcelColor),
            style
                .border()
                .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorder)));
  }

  static Optional<ExcelDifferentialBorder> toExcelDifferentialBorder(
      DifferentialBorderInput border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialBorder(
            border.all().flatMap(WorkbookCommandCellInputConverter::toExcelBorderSide).orElse(null),
            border.top().flatMap(WorkbookCommandCellInputConverter::toExcelBorderSide).orElse(null),
            border
                .right()
                .flatMap(WorkbookCommandCellInputConverter::toExcelBorderSide)
                .orElse(null),
            border
                .bottom()
                .flatMap(WorkbookCommandCellInputConverter::toExcelBorderSide)
                .orElse(null),
            border
                .left()
                .flatMap(WorkbookCommandCellInputConverter::toExcelBorderSide)
                .orElse(null)));
  }

  static ExcelAutofilterFilterColumn toExcelAutofilterFilterColumn(
      AutofilterFilterColumnInput column) {
    return new ExcelAutofilterFilterColumn(
        column.columnId(),
        column.showButton(),
        toExcelAutofilterFilterCriterion(column.criterion()));
  }

  static Optional<ExcelAutofilterSortState> toExcelAutofilterSortState(
      AutofilterSortStateInput sortState) {
    if (sortState == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelAutofilterSortState(
            sortState.range(),
            sortState.caseSensitive(),
            sortState.columnSort(),
            sortState.sortMethod(),
            sortState.conditions().stream()
                .map(WorkbookCommandStructuredInputConverter::toExcelAutofilterSortCondition)
                .toList()));
  }

  private static ExcelAutofilterFilterCriterion toExcelAutofilterFilterCriterion(
      AutofilterFilterCriterionInput criterion) {
    return switch (criterion) {
      case AutofilterFilterCriterionInput.Values values ->
          new ExcelAutofilterFilterCriterion.Values(values.values(), values.includeBlank());
      case AutofilterFilterCriterionInput.Custom custom ->
          new ExcelAutofilterFilterCriterion.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new ExcelAutofilterFilterCriterion.CustomCondition(
                              condition.operator(), condition.value()))
                  .toList());
      case AutofilterFilterCriterionInput.Dynamic dynamic ->
          new ExcelAutofilterFilterCriterion.Dynamic(
              dynamic.type(), dynamic.value().orElse(null), dynamic.maxValue().orElse(null));
      case AutofilterFilterCriterionInput.Top10 top10 ->
          new ExcelAutofilterFilterCriterion.Top10(top10.value(), top10.top(), top10.percent());
      case AutofilterFilterCriterionInput.Color color ->
          new ExcelAutofilterFilterCriterion.Color(
              color.cellColor(),
              WorkbookCommandCellInputConverter.toRequiredExcelColor(
                  color.color(), "autofilter color"));
      case AutofilterFilterCriterionInput.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static ExcelAutofilterSortCondition toExcelAutofilterSortCondition(
      AutofilterSortConditionInput condition) {
    return switch (condition) {
      case AutofilterSortConditionInput.Value value ->
          new ExcelAutofilterSortCondition.Value(value.range(), value.descending());
      case AutofilterSortConditionInput.CellColor cellColor ->
          new ExcelAutofilterSortCondition.CellColor(
              cellColor.range(),
              cellColor.descending(),
              WorkbookCommandCellInputConverter.toRequiredExcelColor(
                  cellColor.color(), "autofilter cell sort color"));
      case AutofilterSortConditionInput.FontColor fontColor ->
          new ExcelAutofilterSortCondition.FontColor(
              fontColor.range(),
              fontColor.descending(),
              WorkbookCommandCellInputConverter.toRequiredExcelColor(
                  fontColor.color(), "autofilter font sort color"));
      case AutofilterSortConditionInput.Icon icon ->
          new ExcelAutofilterSortCondition.Icon(icon.range(), icon.descending(), icon.iconId());
    };
  }

  private static dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold
      toExcelConditionalFormattingThreshold(ConditionalFormattingThresholdInput threshold) {
    return switch (threshold) {
      case ConditionalFormattingThresholdInput.Min _ ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType.MIN,
              null,
              null);
      case ConditionalFormattingThresholdInput.Max _ ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType.MAX,
              null,
              null);
      case ConditionalFormattingThresholdInput.Numeric number ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType.NUMBER,
              null,
              number.value());
      case ConditionalFormattingThresholdInput.Percent percent ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType.PERCENT,
              null,
              percent.value());
      case ConditionalFormattingThresholdInput.Percentile percentile ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType
                  .PERCENTILE,
              null,
              percentile.value());
      case ConditionalFormattingThresholdInput.Formula formula ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType.FORMULA,
              formula.formula(),
              null);
    };
  }
}
