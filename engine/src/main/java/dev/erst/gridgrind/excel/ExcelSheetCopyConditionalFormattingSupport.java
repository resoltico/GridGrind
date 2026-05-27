package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.formula.FormulaType;
import org.jspecify.annotations.Nullable;

/** Copies conditional-formatting blocks and rewrites sheet-local formulas for the cloned sheet. */
final class ExcelSheetCopyConditionalFormattingSupport {
  private ExcelSheetCopyConditionalFormattingSupport() {}

  static void replaceConditionalFormatting(
      List<ExcelConditionalFormattingBlockDefinition> blocks,
      ExcelSheet targetSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    targetSheet.metadata().clearConditionalFormatting(new ExcelRangeSelection.All());
    for (ExcelConditionalFormattingBlockDefinition block : blocks) {
      targetSheet
          .metadata()
          .setConditionalFormatting(
              retargetConditionalFormattingBlock(workbook, block, sourceSheetName, newSheetName));
    }
  }

  static List<ExcelConditionalFormattingBlockDefinition> supportedConditionalFormatting(
      List<ExcelConditionalFormattingBlockSnapshot> blocks, String sourceSheetName) {
    Objects.requireNonNull(blocks, "blocks must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    List<ExcelConditionalFormattingBlockDefinition> copyableBlocks = new ArrayList<>();
    for (ExcelConditionalFormattingBlockSnapshot block : blocks) {
      copyableBlocks.add(
          new ExcelConditionalFormattingBlockDefinition(
              block.ranges(),
              List.copyOf(copyableConditionalFormattingRules(block.rules(), sourceSheetName))));
    }
    return List.copyOf(copyableBlocks);
  }

  static ExcelConditionalFormattingRule copyableRule(
      ExcelConditionalFormattingRuleSnapshot rule, String sourceSheetName) {
    Objects.requireNonNull(rule, "rule must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    return switch (rule) {
      case ExcelConditionalFormattingRuleSnapshot.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              copyableStyle(formulaRule.style(), sourceSheetName));
      case ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              Optional.ofNullable(cellValueRule.formula2()),
              cellValueRule.stopIfTrue(),
              copyableStyle(cellValueRule.style(), sourceSheetName));
      case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(ExcelSheetCopyConditionalFormattingSupport::copyableThreshold)
                  .toList(),
              colorScaleRule.colors().stream()
                  .map(color -> (ExcelColor) ExcelColor.rgb(color))
                  .toList(),
              colorScaleRule.stopIfTrue());
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              ExcelColor.rgb(dataBarRule.color()),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              copyableThreshold(dataBarRule.minThreshold()),
              copyableThreshold(dataBarRule.maxThreshold()),
              dataBarRule.stopIfTrue());
      case ExcelConditionalFormattingRuleSnapshot.IconSetRule iconSetRule ->
          new ExcelConditionalFormattingRule.IconSetRule(
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(ExcelSheetCopyConditionalFormattingSupport::copyableThreshold)
                  .toList(),
              iconSetRule.stopIfTrue());
      case ExcelConditionalFormattingRuleSnapshot.Top10Rule top10Rule ->
          new ExcelConditionalFormattingRule.Top10Rule(
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              top10Rule.stopIfTrue(),
              copyableStyle(top10Rule.style(), sourceSheetName));
      case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
          throw new IllegalArgumentException(
              "cannot copy sheet '"
                  + sourceSheetName
                  + "': unsupported conditional-formatting rule '"
                  + unsupportedRule.kind()
                  + "' is not copyable");
    };
  }

  static Optional<ExcelDifferentialStyle> copyableStyle(
      @Nullable ExcelDifferentialStyleSnapshot style, String sourceSheetName) {
    ExcelWorkbookSheetSupport.requireSheetName(sourceSheetName, "sourceSheetName");
    if (style == null) {
      return Optional.empty();
    }
    if (!style.unsupportedFeatures().isEmpty()) {
      throw new IllegalArgumentException(
          "cannot copy sheet '"
              + sourceSheetName
              + "': conditional-formatting rules with unsupported differential-style features are"
              + " not copyable");
    }
    return Optional.of(
        new ExcelDifferentialStyle(
            Optional.ofNullable(style.numberFormat()),
            Optional.ofNullable(style.bold()),
            Optional.ofNullable(style.italic()),
            Optional.ofNullable(style.fontHeight()),
            Optional.ofNullable(style.fontColor()),
            Optional.ofNullable(style.underline()),
            Optional.ofNullable(style.strikeout()),
            Optional.ofNullable(style.fillColor()),
            Optional.ofNullable(style.border())));
  }

  private static List<ExcelConditionalFormattingRule> copyableConditionalFormattingRules(
      List<ExcelConditionalFormattingRuleSnapshot> rules, String sourceSheetName) {
    List<ExcelConditionalFormattingRule> copyableRules = new ArrayList<>();
    for (ExcelConditionalFormattingRuleSnapshot rule : rules) {
      copyableRules.add(copyableRule(rule, sourceSheetName));
    }
    return List.copyOf(copyableRules);
  }

  private static ExcelConditionalFormattingThreshold copyableThreshold(
      ExcelConditionalFormattingThresholdSnapshot threshold) {
    return new ExcelConditionalFormattingThreshold(
        threshold.type(), threshold.formula(), threshold.value());
  }

  private static ExcelConditionalFormattingBlockDefinition retargetConditionalFormattingBlock(
      ExcelWorkbook workbook,
      ExcelConditionalFormattingBlockDefinition block,
      String sourceSheetName,
      String newSheetName) {
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(newSheetName);
    return new ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
            .map(
                rule ->
                    retargetConditionalFormattingRule(
                        workbook, rule, targetSheetIndex, sourceSheetName, newSheetName))
            .toList());
  }

  private static ExcelConditionalFormattingRule retargetConditionalFormattingRule(
      ExcelWorkbook workbook,
      ExcelConditionalFormattingRule rule,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    return switch (rule) {
      case ExcelConditionalFormattingRule.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              ExcelSheetCopySupport.retargetFormula(
                  workbook,
                  formulaRule.formula(),
                  FormulaType.CONDFORMAT,
                  targetSheetIndex,
                  sourceSheetName,
                  newSheetName),
              formulaRule.stopIfTrue(),
              formulaRule.style());
      case ExcelConditionalFormattingRule.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              ExcelSheetCopySupport.retargetFormula(
                  workbook,
                  cellValueRule.formula1(),
                  FormulaType.CONDFORMAT,
                  targetSheetIndex,
                  sourceSheetName,
                  newSheetName),
              ExcelSheetCopySupport.retargetOptionalFormula(
                  workbook,
                  cellValueRule.formula2(),
                  FormulaType.CONDFORMAT,
                  targetSheetIndex,
                  sourceSheetName,
                  newSheetName),
              cellValueRule.stopIfTrue(),
              cellValueRule.style());
      case ExcelConditionalFormattingRule.ColorScaleRule colorScaleRule -> colorScaleRule;
      case ExcelConditionalFormattingRule.DataBarRule dataBarRule -> dataBarRule;
      case ExcelConditionalFormattingRule.IconSetRule iconSetRule -> iconSetRule;
      case ExcelConditionalFormattingRule.Top10Rule top10Rule -> top10Rule;
    };
  }
}
