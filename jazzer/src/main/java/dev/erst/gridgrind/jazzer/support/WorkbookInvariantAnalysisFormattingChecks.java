package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderReport;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideReport;
import dev.erst.gridgrind.contract.dto.DifferentialStyleReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;

/** Owns invariant checks for validation and conditional-formatting rule payloads. */
final class WorkbookInvariantAnalysisFormattingChecks {
  private WorkbookInvariantAnalysisFormattingChecks() {}

  static void requireDataValidationEntryShape(DataValidationEntryReport validation) {
    WorkbookInvariantChecks.require(
        validation.ranges() != null, "data validation ranges must not be null");
    WorkbookInvariantChecks.require(
        !validation.ranges().isEmpty(), "data validation ranges must not be empty");
    validation
        .ranges()
        .forEach(range -> WorkbookInvariantChecks.requireNonBlank(range, "data validation range"));

    switch (validation) {
      case DataValidationEntryReport.Supported supported ->
          requireSupportedDataValidationShape(supported.validation());
      case DataValidationEntryReport.Unsupported unsupported -> {
        WorkbookInvariantChecks.requireNonBlank(unsupported.kind(), "data validation kind");
        WorkbookInvariantChecks.requireNonBlank(unsupported.detail(), "data validation detail");
      }
    }
  }

  static void requireConditionalFormattingEntryShape(
      dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport conditionalFormatting) {
    WorkbookInvariantChecks.require(
        conditionalFormatting.ranges() != null, "conditional formatting ranges must not be null");
    WorkbookInvariantChecks.require(
        !conditionalFormatting.ranges().isEmpty(),
        "conditional formatting ranges must not be empty");
    conditionalFormatting
        .ranges()
        .forEach(
            range ->
                WorkbookInvariantChecks.requireNonBlank(range, "conditional formatting range"));
    WorkbookInvariantChecks.require(
        conditionalFormatting.rules() != null, "conditional formatting rules must not be null");
    WorkbookInvariantChecks.require(
        !conditionalFormatting.rules().isEmpty(), "conditional formatting rules must not be empty");
    conditionalFormatting
        .rules()
        .forEach(WorkbookInvariantAnalysisFormattingChecks::requireConditionalFormattingRuleShape);
  }

  static void requireConditionalFormattingRuleShape(ConditionalFormattingRuleReport rule) {
    WorkbookInvariantChecks.require(
        rule.priority() > 0, "conditional formatting priority must be greater than 0");
    switch (rule) {
      case ConditionalFormattingRuleReport.FormulaRule formulaRule -> {
        WorkbookInvariantChecks.requireNonBlank(
            formulaRule.formula(), "conditional formatting formula");
        WorkbookInvariantChecks.require(
            formulaRule.style() != null, "conditional formatting style must not be null");
        formulaRule
            .style()
            .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialStyleShape);
      }
      case ConditionalFormattingRuleReport.CellValueRule cellValueRule -> {
        WorkbookInvariantChecks.require(
            cellValueRule.operator() != null, "conditional formatting operator must not be null");
        WorkbookInvariantChecks.requireNonBlank(
            cellValueRule.formula1(), "conditional formatting formula1");
        cellValueRule
            .formula2()
            .ifPresent(
                formula ->
                    WorkbookInvariantChecks.requireNonBlank(
                        formula, "conditional formatting formula2"));
        cellValueRule
            .style()
            .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialStyleShape);
      }
      case ConditionalFormattingRuleReport.ColorScaleRule colorScaleRule -> {
        WorkbookInvariantChecks.require(
            colorScaleRule.thresholds() != null,
            "conditional formatting thresholds must not be null");
        WorkbookInvariantChecks.require(
            !colorScaleRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        colorScaleRule
            .thresholds()
            .forEach(
                WorkbookInvariantAnalysisFormattingChecks
                    ::requireConditionalFormattingThresholdShape);
        WorkbookInvariantChecks.require(
            colorScaleRule.colors() != null, "conditional formatting colors must not be null");
        WorkbookInvariantChecks.require(
            !colorScaleRule.colors().isEmpty(), "conditional formatting colors must not be empty");
        colorScaleRule
            .colors()
            .forEach(
                color ->
                    WorkbookInvariantChecks.requireNonBlank(color, "conditional formatting color"));
      }
      case ConditionalFormattingRuleReport.DataBarRule dataBarRule -> {
        WorkbookInvariantChecks.requireNonBlank(
            dataBarRule.color(), "conditional formatting color");
        requireConditionalFormattingThresholdShape(dataBarRule.minThreshold());
        requireConditionalFormattingThresholdShape(dataBarRule.maxThreshold());
        WorkbookInvariantChecks.require(
            dataBarRule.widthMin() >= 0, "conditional formatting widthMin must not be negative");
        WorkbookInvariantChecks.require(
            dataBarRule.widthMax() >= 0, "conditional formatting widthMax must not be negative");
      }
      case ConditionalFormattingRuleReport.IconSetRule iconSetRule -> {
        WorkbookInvariantChecks.require(
            iconSetRule.iconSet() != null, "conditional formatting iconSet must not be null");
        WorkbookInvariantChecks.require(
            iconSetRule.thresholds() != null, "conditional formatting thresholds must not be null");
        WorkbookInvariantChecks.require(
            !iconSetRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        iconSetRule
            .thresholds()
            .forEach(
                WorkbookInvariantAnalysisFormattingChecks
                    ::requireConditionalFormattingThresholdShape);
      }
      case ConditionalFormattingRuleReport.Top10Rule top10Rule -> {
        WorkbookInvariantChecks.require(
            top10Rule.rank() >= 0, "conditional formatting rank must not be negative");
        top10Rule
            .style()
            .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialStyleShape);
      }
      case ConditionalFormattingRuleReport.UnsupportedRule unsupportedRule -> {
        WorkbookInvariantChecks.requireNonBlank(
            unsupportedRule.kind(), "conditional formatting kind");
        WorkbookInvariantChecks.requireNonBlank(
            unsupportedRule.detail(), "conditional formatting detail");
      }
    }
  }

  static void requireConditionalFormattingThresholdShape(
      ConditionalFormattingThresholdReport threshold) {
    WorkbookInvariantChecks.require(
        threshold != null, "conditional formatting threshold must not be null");
    WorkbookInvariantChecks.require(
        threshold.type() != null, "conditional formatting threshold type must not be null");
  }

  static void requireDifferentialStyleShape(DifferentialStyleReport style) {
    WorkbookInvariantChecks.require(style != null, "conditional formatting style must not be null");
    if (style.numberFormat() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          style.numberFormat(), "conditional formatting numberFormat");
    }
    if (style.fontHeight() != null) {
      WorkbookInvariantCellSurfaceChecks.requireFontHeightShape(style.fontHeight());
    }
    style
        .fontColor()
        .ifPresent(
            color ->
                WorkbookInvariantCellStyleChecks.requireCellColorShape(
                    color, "conditional formatting fontColor"));
    style
        .fillColor()
        .ifPresent(
            color ->
                WorkbookInvariantCellStyleChecks.requireCellColorShape(
                    color, "conditional formatting fillColor"));
    style
        .border()
        .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialBorderShape);
    WorkbookInvariantChecks.require(
        style.unsupportedFeatures() != null,
        "conditional formatting unsupportedFeatures must not be null");
    style
        .unsupportedFeatures()
        .forEach(
            feature ->
                WorkbookInvariantChecks.require(
                    feature != null,
                    "conditional formatting unsupported feature must not be null"));
  }

  static void requireDifferentialBorderShape(DifferentialBorderReport border) {
    WorkbookInvariantChecks.require(
        border != null, "conditional formatting border must not be null");
    border
        .all()
        .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialBorderSideShape);
    border
        .top()
        .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialBorderSideShape);
    border
        .right()
        .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialBorderSideShape);
    border
        .bottom()
        .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialBorderSideShape);
    border
        .left()
        .ifPresent(WorkbookInvariantAnalysisFormattingChecks::requireDifferentialBorderSideShape);
  }

  static void requireDifferentialBorderSideShape(DifferentialBorderSideReport side) {
    WorkbookInvariantChecks.require(
        side != null, "conditional formatting border side must not be null");
    WorkbookInvariantChecks.require(
        side.style() != null, "conditional formatting border style must not be null");
    side.color()
        .ifPresent(
            color ->
                WorkbookInvariantCellStyleChecks.requireCellColorShape(
                    color, "conditional formatting border color"));
  }

  static void requireTableStyleShape(TableStyleReport style) {
    switch (style) {
      case TableStyleReport.None _ -> {}
      case TableStyleReport.Named named ->
          WorkbookInvariantChecks.requireNonBlank(named.name(), "table style name");
    }
  }

  static void requireSupportedDataValidationShape(
      DataValidationEntryReport.DataValidationDefinitionReport validation) {
    WorkbookInvariantChecks.require(
        validation != null, "data validation definition must not be null");
    WorkbookInvariantChecks.require(
        validation.rule() != null, "data validation rule must not be null");
    switch (validation.rule()) {
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.ExplicitList explicitList -> {
        WorkbookInvariantChecks.require(
            explicitList.values() != null, "explicit list values must not be null");
        explicitList
            .values()
            .forEach(
                value -> WorkbookInvariantChecks.requireNonBlank(value, "explicit list value"));
      }
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.FormulaList formulaList ->
          WorkbookInvariantChecks.requireNonBlank(formulaList.formula(), "formula list formula");
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.WholeNumber wholeNumber ->
          requireComparisonRuleShape(wholeNumber.operator(), wholeNumber.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.DecimalNumber decimalNumber ->
          requireComparisonRuleShape(decimalNumber.operator(), decimalNumber.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.DateRule dateRule ->
          requireComparisonRuleShape(dateRule.operator(), dateRule.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.TimeRule timeRule ->
          requireComparisonRuleShape(timeRule.operator(), timeRule.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.TextLength textLength ->
          requireComparisonRuleShape(textLength.operator(), textLength.formula1());
      case dev.erst.gridgrind.contract.dto.DataValidationRuleInput.CustomFormula customFormula ->
          WorkbookInvariantChecks.requireNonBlank(
              customFormula.formula(), "custom validation formula");
    }
    validation
        .prompt()
        .ifPresent(
            prompt -> {
              WorkbookInvariantChecks.requireNonBlank(
                  prompt.title(), "data validation prompt title");
              WorkbookInvariantChecks.requireNonBlank(prompt.text(), "data validation prompt text");
            });
    validation
        .errorAlert()
        .ifPresent(
            errorAlert -> {
              WorkbookInvariantChecks.require(
                  errorAlert.style() != null, "data validation error style must not be null");
              WorkbookInvariantChecks.requireNonBlank(
                  errorAlert.title(), "data validation error title");
              WorkbookInvariantChecks.requireNonBlank(
                  errorAlert.text(), "data validation error text");
            });
  }

  static void requireComparisonRuleShape(Object operator, String formula1) {
    WorkbookInvariantChecks.require(operator != null, "comparison operator must not be null");
    WorkbookInvariantChecks.requireNonBlank(formula1, "comparison formula1");
  }
}
