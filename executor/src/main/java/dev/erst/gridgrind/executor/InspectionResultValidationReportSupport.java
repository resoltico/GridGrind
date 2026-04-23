package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideReport;
import dev.erst.gridgrind.contract.dto.DifferentialStyleReport;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;

/** Converts validation and conditional-format workbook snapshots into protocol reports. */
final class InspectionResultValidationReportSupport {
  private InspectionResultValidationReportSupport() {}

  static DataValidationEntryReport toDataValidationEntryReport(
      ExcelDataValidationSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelDataValidationSnapshot.Supported supported ->
          new DataValidationEntryReport.Supported(
              supported.ranges(), toDataValidationDefinitionReport(supported.validation()));
      case ExcelDataValidationSnapshot.Unsupported unsupported ->
          new DataValidationEntryReport.Unsupported(
              unsupported.ranges(), unsupported.kind(), unsupported.detail());
    };
  }

  static DataValidationEntryReport.DataValidationDefinitionReport toDataValidationDefinitionReport(
      ExcelDataValidationDefinition definition) {
    return new DataValidationEntryReport.DataValidationDefinitionReport(
        toDataValidationRuleInput(definition.rule()),
        definition.allowBlank(),
        definition.suppressDropDownArrow(),
        definition.prompt() == null
            ? null
            : new DataValidationPromptInput(
                new TextSourceInput.Inline(definition.prompt().title()),
                new TextSourceInput.Inline(definition.prompt().text()),
                definition.prompt().showPromptBox()),
        definition.errorAlert() == null
            ? null
            : new DataValidationErrorAlertInput(
                definition.errorAlert().style(),
                new TextSourceInput.Inline(definition.errorAlert().title()),
                new TextSourceInput.Inline(definition.errorAlert().text()),
                definition.errorAlert().showErrorBox()));
  }

  static ConditionalFormattingEntryReport toConditionalFormattingEntryReport(
      ExcelConditionalFormattingBlockSnapshot block) {
    return new ConditionalFormattingEntryReport(
        block.ranges(),
        block.rules().stream()
            .map(InspectionResultValidationReportSupport::toConditionalFormattingRuleReport)
            .toList());
  }

  static ConditionalFormattingRuleReport toConditionalFormattingRuleReport(
      ExcelConditionalFormattingRuleSnapshot rule) {
    return switch (rule) {
      case ExcelConditionalFormattingRuleSnapshot.FormulaRule formulaRule ->
          new ConditionalFormattingRuleReport.FormulaRule(
              formulaRule.priority(),
              formulaRule.stopIfTrue(),
              formulaRule.formula(),
              toDifferentialStyleReport(formulaRule.style()));
      case ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule ->
          new ConditionalFormattingRuleReport.CellValueRule(
              cellValueRule.priority(),
              cellValueRule.stopIfTrue(),
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              toDifferentialStyleReport(cellValueRule.style()));
      case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule colorScaleRule ->
          new ConditionalFormattingRuleReport.ColorScaleRule(
              colorScaleRule.priority(),
              colorScaleRule.stopIfTrue(),
              colorScaleRule.thresholds().stream()
                  .map(
                      InspectionResultValidationReportSupport
                          ::toConditionalFormattingThresholdReport)
                  .toList(),
              colorScaleRule.colors());
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule dataBarRule ->
          new ConditionalFormattingRuleReport.DataBarRule(
              dataBarRule.priority(),
              dataBarRule.stopIfTrue(),
              dataBarRule.color(),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toConditionalFormattingThresholdReport(dataBarRule.minThreshold()),
              toConditionalFormattingThresholdReport(dataBarRule.maxThreshold()));
      case ExcelConditionalFormattingRuleSnapshot.IconSetRule iconSetRule ->
          new ConditionalFormattingRuleReport.IconSetRule(
              iconSetRule.priority(),
              iconSetRule.stopIfTrue(),
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(
                      InspectionResultValidationReportSupport
                          ::toConditionalFormattingThresholdReport)
                  .toList());
      case ExcelConditionalFormattingRuleSnapshot.Top10Rule top10Rule ->
          new ConditionalFormattingRuleReport.Top10Rule(
              top10Rule.priority(),
              top10Rule.stopIfTrue(),
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              toDifferentialStyleReport(top10Rule.style()));
      case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
          new ConditionalFormattingRuleReport.UnsupportedRule(
              unsupportedRule.priority(),
              unsupportedRule.stopIfTrue(),
              unsupportedRule.kind(),
              unsupportedRule.detail());
    };
  }

  static ConditionalFormattingThresholdReport toConditionalFormattingThresholdReport(
      ExcelConditionalFormattingThresholdSnapshot threshold) {
    return new ConditionalFormattingThresholdReport(
        threshold.type(), threshold.formula(), threshold.value());
  }

  static DifferentialStyleReport toDifferentialStyleReport(ExcelDifferentialStyleSnapshot style) {
    if (style == null) {
      return null;
    }
    return new DifferentialStyleReport(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        InspectionResultCellReportSupport.toFontHeightReport(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toDifferentialBorderReport(style.border()),
        style.unsupportedFeatures());
  }

  static DifferentialBorderReport toDifferentialBorderReport(ExcelDifferentialBorder border) {
    if (border == null) {
      return null;
    }
    return new DifferentialBorderReport(
        toDifferentialBorderSideReport(border.all()),
        toDifferentialBorderSideReport(border.top()),
        toDifferentialBorderSideReport(border.right()),
        toDifferentialBorderSideReport(border.bottom()),
        toDifferentialBorderSideReport(border.left()));
  }

  static DifferentialBorderSideReport toDifferentialBorderSideReport(
      ExcelDifferentialBorderSide side) {
    return side == null ? null : new DifferentialBorderSideReport(side.style(), side.color());
  }

  private static DataValidationRuleInput toDataValidationRuleInput(ExcelDataValidationRule rule) {
    return switch (rule) {
      case ExcelDataValidationRule.ExplicitList explicitList ->
          new DataValidationRuleInput.ExplicitList(explicitList.values());
      case ExcelDataValidationRule.FormulaList formulaList ->
          new DataValidationRuleInput.FormulaList(formulaList.formula());
      case ExcelDataValidationRule.WholeNumber wholeNumber ->
          new DataValidationRuleInput.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case ExcelDataValidationRule.DecimalNumber decimalNumber ->
          new DataValidationRuleInput.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case ExcelDataValidationRule.DateRule dateRule ->
          new DataValidationRuleInput.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case ExcelDataValidationRule.TimeRule timeRule ->
          new DataValidationRuleInput.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case ExcelDataValidationRule.TextLength textLength ->
          new DataValidationRuleInput.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case ExcelDataValidationRule.CustomFormula customFormula ->
          new DataValidationRuleInput.CustomFormula(customFormula.formula());
    };
  }
}
