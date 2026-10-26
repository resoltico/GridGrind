package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
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
      DataValidationSpec validation) {
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

  static ExcelDataValidationRule toExcelDataValidationRule(DataValidationRuleSpec rule) {
    return switch (rule) {
      case DataValidationRuleSpec.ExplicitList explicitList ->
          new ExcelDataValidationRule.ExplicitList(explicitList.values());
      case DataValidationRuleSpec.FormulaList formulaList ->
          new ExcelDataValidationRule.FormulaList(formulaList.formula());
      case DataValidationRuleSpec.WholeNumber wholeNumber ->
          new ExcelDataValidationRule.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case DataValidationRuleSpec.DecimalNumber decimalNumber ->
          new ExcelDataValidationRule.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case DataValidationRuleSpec.DateRule dateRule ->
          new ExcelDataValidationRule.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case DataValidationRuleSpec.TimeRule timeRule ->
          new ExcelDataValidationRule.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case DataValidationRuleSpec.TextLength textLength ->
          new ExcelDataValidationRule.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case DataValidationRuleSpec.CustomFormula customFormula ->
          new ExcelDataValidationRule.CustomFormula(customFormula.formula());
    };
  }

  static Optional<ExcelDataValidationPrompt> toExcelDataValidationPrompt(
      DataValidationPromptSpec prompt) {
    return prompt == null
        ? Optional.empty()
        : Optional.of(
            new ExcelDataValidationPrompt(
                WorkbookCommandSourceSupport.inlineText(prompt.title(), "validation prompt title"),
                WorkbookCommandSourceSupport.inlineText(prompt.text(), "validation prompt text"),
                prompt.showPromptBox()));
  }

  static Optional<ExcelDataValidationErrorAlert> toExcelDataValidationErrorAlert(
      DataValidationErrorAlertSpec errorAlert) {
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
      List<String> ranges, ConditionalFormattingSpec definition) {
    return new ExcelConditionalFormattingBlockDefinition(
        ranges,
        definition.rules().stream()
            .map(WorkbookCommandStructuredInputConverter::toExcelConditionalFormattingRule)
            .toList());
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleSpec rule) {
    return switch (rule) {
      case ConditionalFormattingRuleSpec.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              formulaRule
                  .style()
                  .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialStyle));
      case ConditionalFormattingRuleSpec.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              cellValueRule
                  .style()
                  .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialStyle));
      case ConditionalFormattingRuleSpec.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(
                      WorkbookCommandStructuredInputConverter
                          ::toExcelConditionalFormattingThreshold)
                  .toList(),
              colorScaleRule.colors().stream().map(color -> color).toList(),
              colorScaleRule.stopIfTrue());
      case ConditionalFormattingRuleSpec.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              dataBarRule.color(),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toExcelConditionalFormattingThreshold(dataBarRule.minThreshold()),
              toExcelConditionalFormattingThreshold(dataBarRule.maxThreshold()),
              dataBarRule.stopIfTrue());
      case ConditionalFormattingRuleSpec.IconSetRule iconSetRule ->
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
      case ConditionalFormattingRuleSpec.Top10Rule top10Rule ->
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

  static Optional<ExcelDifferentialStyle> toExcelDifferentialStyle(DifferentialStylePatch style) {
    if (style == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialStyle(
            style.numberFormat(),
            style.bold(),
            style.italic(),
            style.fontHeight().flatMap(WorkbookCommandCellValueSpecConverter::toFontHeightReport),
            style.fontColor(),
            style.underline(),
            style.strikeout(),
            style.fillColor(),
            style
                .border()
                .flatMap(WorkbookCommandStructuredInputConverter::toExcelDifferentialBorder)));
  }

  static Optional<ExcelDifferentialBorder> toExcelDifferentialBorder(
      DifferentialBorderPatch border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialBorder(
            border.all().orElse(null),
            border.top().orElse(null),
            border.right().orElse(null),
            border.bottom().orElse(null),
            border.left().orElse(null)));
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
              dynamic.type(), dynamic.value().orElse(null), dynamic.maxValue().orElse(null));
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

  private static dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold
      toExcelConditionalFormattingThreshold(ConditionalFormattingThresholdSpec threshold) {
    return switch (threshold) {
      case ConditionalFormattingThresholdSpec.Min _ ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType.MIN,
              null,
              null);
      case ConditionalFormattingThresholdSpec.Max _ ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType.MAX,
              null,
              null);
      case ConditionalFormattingThresholdSpec.Numeric number ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType.NUMBER,
              null,
              number.value());
      case ConditionalFormattingThresholdSpec.Percent percent ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType.PERCENT,
              null,
              percent.value());
      case ConditionalFormattingThresholdSpec.Percentile percentile ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType.PERCENTILE,
              null,
              percentile.value());
      case ConditionalFormattingThresholdSpec.Formula formula ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold(
              dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType.FORMULA,
              formula.formula(),
              null);
    };
  }
}
