package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
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
        toExcelDataValidationPrompt(validation.prompt()).orElse(null),
        toExcelDataValidationErrorAlert(validation.errorAlert()).orElse(null));
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
      ConditionalFormattingBlockInput block) {
    return new ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
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
              toExcelDifferentialStyle(formulaRule.style()).orElse(null));
      case ConditionalFormattingRuleInput.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              toExcelDifferentialStyle(cellValueRule.style()).orElse(null));
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
              toExcelDifferentialStyle(top10Rule.style()).orElse(null));
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
            WorkbookCommandCellInputConverter.toExcelFontHeight(style.fontHeight()).orElse(null),
            style.fontColor().orElse(null),
            style.underline(),
            style.strikeout(),
            style.fillColor().orElse(null),
            style
                .border()
                .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorder)
                .orElse(null)));
  }

  static Optional<ExcelDifferentialBorder> toExcelDifferentialBorder(
      DifferentialBorderInput border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialBorder(
            border
                .all()
                .map(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorderSide)
                .orElse(null),
            border
                .top()
                .map(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorderSide)
                .orElse(null),
            border
                .right()
                .map(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorderSide)
                .orElse(null),
            border
                .bottom()
                .map(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorderSide)
                .orElse(null),
            border
                .left()
                .map(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorderSide)
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

  private static ExcelConditionalFormattingThreshold toExcelConditionalFormattingThreshold(
      ConditionalFormattingThresholdInput threshold) {
    return new ExcelConditionalFormattingThreshold(
        threshold.type(), threshold.formula(), threshold.value());
  }

  private static ExcelDifferentialBorderSide toExcelDifferentialBorderSide(
      DifferentialBorderSideInput side) {
    return new ExcelDifferentialBorderSide(side.style(), side.color().orElse(null));
  }
}
