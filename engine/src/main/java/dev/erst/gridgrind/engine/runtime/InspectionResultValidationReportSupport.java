package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderSideReport;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
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
      ExcelDataValidationDefinition definition) {
    return new DataValidationEntryReport.DataValidationDefinitionReport(
        toDataValidationRuleInput(definition.rule()),
        definition.allowBlank(),
        definition.suppressDropDownArrow(),
        promptInput(definition),
        errorAlertInput(definition));
  }

  static ConditionalFormattingEntryReport toConditionalFormattingEntryReport(
      ExcelConditionalFormattingBlockSnapshot block) {
    return new ConditionalFormattingEntryReport(
        block.ranges().stream().map(RangeLabel::new).toList(),
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
              Optional.ofNullable(cellValueRule.formula2()),
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
              colorScaleRule.colors().stream().map(RgbHex::new).toList());
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule dataBarRule ->
          new ConditionalFormattingRuleReport.DataBarRule(
              dataBarRule.priority(),
              dataBarRule.stopIfTrue(),
              java.util.Optional.ofNullable(dataBarRule.color())
                  .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new),
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

  static Optional<DifferentialStyleReport> toDifferentialStyleReport(
      @org.jspecify.annotations.Nullable ExcelDifferentialStyleSnapshot style) {
    if (style == null) {
      return Optional.empty();
    }
    return Optional.of(
        new DifferentialStyleReport(
            style.numberFormat(),
            style.bold(),
            style.italic(),
            style.fontHeight(),
            InspectionResultCellStyleReportSupport.toCellColorReport(style.fontColor()),
            style.underline(),
            style.strikeout(),
            InspectionResultCellStyleReportSupport.toCellColorReport(style.fillColor()),
            toDifferentialBorderReport(style.border()),
            style.unsupportedFeatures()));
  }

  static Optional<DifferentialBorderReport> toDifferentialBorderReport(
      @org.jspecify.annotations.Nullable ExcelDifferentialBorder border) {
    if (border == null) {
      return Optional.empty();
    }
    return Optional.of(
        new DifferentialBorderReport(
            toDifferentialBorderSideReport(border.all()),
            toDifferentialBorderSideReport(border.top()),
            toDifferentialBorderSideReport(border.right()),
            toDifferentialBorderSideReport(border.bottom()),
            toDifferentialBorderSideReport(border.left())));
  }

  static Optional<DifferentialBorderSideReport> toDifferentialBorderSideReport(
      @org.jspecify.annotations.Nullable BorderSidePatch side) {
    return side == null
        ? Optional.empty()
        : Optional.of(
            new DifferentialBorderSideReport(
                side.style(),
                side.color().flatMap(InspectionResultCellStyleReportSupport::toCellColorReport)));
  }

  private static Optional<DataValidationPromptSpec> promptInput(
      ExcelDataValidationDefinition definition) {
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
      ExcelDataValidationDefinition definition) {
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

  private static DataValidationRuleSpec toDataValidationRuleInput(ExcelDataValidationRule rule) {
    return switch (rule) {
      case ExcelDataValidationRule.ExplicitList explicitList ->
          new DataValidationRuleSpec.ExplicitList(explicitList.values());
      case ExcelDataValidationRule.FormulaList formulaList ->
          new DataValidationRuleSpec.FormulaList(formulaList.formula());
      case ExcelDataValidationRule.WholeNumber wholeNumber ->
          new DataValidationRuleSpec.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case ExcelDataValidationRule.DecimalNumber decimalNumber ->
          new DataValidationRuleSpec.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case ExcelDataValidationRule.DateRule dateRule ->
          new DataValidationRuleSpec.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case ExcelDataValidationRule.TimeRule timeRule ->
          new DataValidationRuleSpec.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case ExcelDataValidationRule.TextLength textLength ->
          new DataValidationRuleSpec.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case ExcelDataValidationRule.CustomFormula customFormula ->
          new DataValidationRuleSpec.CustomFormula(customFormula.formula());
    };
  }
}
