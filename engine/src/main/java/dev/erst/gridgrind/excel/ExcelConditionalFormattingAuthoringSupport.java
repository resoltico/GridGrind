package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingThreshold;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;

/** Applies authored conditional-formatting blocks onto one sheet. */
final class ExcelConditionalFormattingAuthoringSupport {
  void setConditionalFormatting(XSSFSheet sheet, ExcelConditionalFormattingBlockDefinition block) {
    List<ExcelRange> targetRanges = block.ranges().stream().map(ExcelRange::parse).toList();
    removeIntersectingBlocks(sheet, targetRanges);

    XSSFConditionalFormattingRule[] rules =
        block.rules().stream()
            .map(rule -> createRule(sheet, rule))
            .toArray(XSSFConditionalFormattingRule[]::new);
    int blockIndex =
        sheet
            .getSheetConditionalFormatting()
            .addConditionalFormatting(
                targetRanges.stream()
                    .map(ExcelSheetStructureSupport::toCellRangeAddress)
                    .toArray(CellRangeAddress[]::new),
                rules);

    CTConditionalFormatting ctBlock =
        ExcelConditionalFormattingSnapshotSupport.ctBlock(sheet, blockIndex);
    for (int ruleIndex = 0; ruleIndex < block.rules().size(); ruleIndex++) {
      CTCfRule ctRule = ctBlock.getCfRuleArray(ruleIndex);
      switch (block.rules().get(ruleIndex)) {
        case ExcelConditionalFormattingRule.FormulaRule formulaRule -> {
          ctRule.setStopIfTrue(formulaRule.stopIfTrue());
          if (formulaRule.style() != null) {
            ExcelConditionalFormattingStyleSupport.applyStyle(
                sheet.getWorkbook(), ctRule, formulaRule.style());
          }
        }
        case ExcelConditionalFormattingRule.CellValueRule cellValueRule -> {
          ctRule.setStopIfTrue(cellValueRule.stopIfTrue());
          if (cellValueRule.style() != null) {
            ExcelConditionalFormattingStyleSupport.applyStyle(
                sheet.getWorkbook(), ctRule, cellValueRule.style());
          }
        }
        case ExcelConditionalFormattingRule.ColorScaleRule colorScaleRule ->
            ctRule.setStopIfTrue(colorScaleRule.stopIfTrue());
        case ExcelConditionalFormattingRule.DataBarRule dataBarRule ->
            ctRule.setStopIfTrue(dataBarRule.stopIfTrue());
        case ExcelConditionalFormattingRule.IconSetRule iconSetRule ->
            ctRule.setStopIfTrue(iconSetRule.stopIfTrue());
        case ExcelConditionalFormattingRule.Top10Rule top10Rule -> {
          ctRule.setStopIfTrue(top10Rule.stopIfTrue());
          if (top10Rule.style() != null) {
            ExcelConditionalFormattingStyleSupport.applyStyle(
                sheet.getWorkbook(), ctRule, top10Rule.style());
          }
          applyTop10Rule(ctRule, top10Rule);
        }
      }
    }
    normalizePriorities(sheet);
  }

  void clearConditionalFormatting(XSSFSheet sheet, ExcelRangeSelection selection) {
    switch (selection) {
      case ExcelRangeSelection.All _ -> clearAll(sheet);
      case ExcelRangeSelection.Selected selected ->
          removeIntersectingBlocks(
              sheet, selected.ranges().stream().map(ExcelRange::parse).toList());
    }
    normalizePriorities(sheet);
  }

  private XSSFConditionalFormattingRule createRule(
      XSSFSheet sheet, ExcelConditionalFormattingRule rule) {
    return switch (rule) {
      case ExcelConditionalFormattingRule.FormulaRule formulaRule ->
          sheet
              .getSheetConditionalFormatting()
              .createConditionalFormattingRule(formulaRule.formula());
      case ExcelConditionalFormattingRule.CellValueRule cellValueRule ->
          cellValueRule.formula2() == null
              ? sheet
                  .getSheetConditionalFormatting()
                  .createConditionalFormattingRule(
                      ExcelComparisonOperatorPoiBridge.toPoi(cellValueRule.operator()),
                      cellValueRule.formula1())
              : sheet
                  .getSheetConditionalFormatting()
                  .createConditionalFormattingRule(
                      ExcelComparisonOperatorPoiBridge.toPoi(cellValueRule.operator()),
                      cellValueRule.formula1(),
                      cellValueRule.formula2());
      case ExcelConditionalFormattingRule.ColorScaleRule colorScaleRule ->
          createColorScaleRule(sheet, colorScaleRule);
      case ExcelConditionalFormattingRule.DataBarRule dataBarRule ->
          createDataBarRule(sheet, dataBarRule);
      case ExcelConditionalFormattingRule.IconSetRule iconSetRule ->
          createIconSetRule(sheet, iconSetRule);
      case ExcelConditionalFormattingRule.Top10Rule _ ->
          sheet
              .getSheetConditionalFormatting()
              .createConditionalFormattingRule(ComparisonOperator.GT, "0");
    };
  }

  private XSSFConditionalFormattingRule createColorScaleRule(
      XSSFSheet sheet, ExcelConditionalFormattingRule.ColorScaleRule rule) {
    XSSFConditionalFormattingRule created =
        sheet.getSheetConditionalFormatting().createConditionalFormattingColorScaleRule();
    var formatting = created.getColorScaleFormatting();
    formatting.setNumControlPoints(rule.thresholds().size());
    XSSFConditionalFormattingThreshold[] thresholds = formatting.getThresholds();
    for (int index = 0; index < thresholds.length; index++) {
      applyThreshold(thresholds[index], rule.thresholds().get(index));
    }
    formatting.setColors(
        rule.colors().stream()
            .map(color -> ExcelColorSupport.toXssfColor(sheet.getWorkbook(), color))
            .toArray(XSSFColor[]::new));
    return created;
  }

  private XSSFConditionalFormattingRule createDataBarRule(
      XSSFSheet sheet, ExcelConditionalFormattingRule.DataBarRule rule) {
    XSSFConditionalFormattingRule created =
        sheet
            .getSheetConditionalFormatting()
            .createConditionalFormattingRule(
                ExcelColorSupport.toXssfColor(sheet.getWorkbook(), rule.color()));
    var formatting = created.getDataBarFormatting();
    formatting.setIconOnly(rule.iconOnly());
    formatting.setWidthMin(rule.widthMin());
    formatting.setWidthMax(rule.widthMax());
    applyThreshold(formatting.getMinThreshold(), rule.minThreshold());
    applyThreshold(formatting.getMaxThreshold(), rule.maxThreshold());
    return created;
  }

  private XSSFConditionalFormattingRule createIconSetRule(
      XSSFSheet sheet, ExcelConditionalFormattingRule.IconSetRule rule) {
    XSSFConditionalFormattingRule created =
        sheet
            .getSheetConditionalFormatting()
            .createConditionalFormattingRule(
                ExcelConditionalFormattingPoiBridge.toPoi(rule.iconSet()));
    var formatting = created.getMultiStateFormatting();
    formatting.setIconOnly(rule.iconOnly());
    formatting.setReversed(rule.reversed());
    XSSFConditionalFormattingThreshold[] thresholds = formatting.getThresholds();
    for (int index = 0; index < thresholds.length; index++) {
      applyThreshold(thresholds[index], rule.thresholds().get(index));
    }
    return created;
  }

  private static void applyThreshold(
      ConditionalFormattingThreshold threshold,
      ExcelConditionalFormattingThreshold authoredThreshold) {
    threshold.setRangeType(ExcelConditionalFormattingPoiBridge.toPoi(authoredThreshold.type()));
    if (authoredThreshold.formula() != null) {
      threshold.setFormula(authoredThreshold.formula());
    } else if (authoredThreshold.value() != null) {
      threshold.setValue(authoredThreshold.value());
    }
  }

  private static void applyTop10Rule(
      CTCfRule ctRule, ExcelConditionalFormattingRule.Top10Rule top10Rule) {
    while (ctRule.sizeOfFormulaArray() > 0) {
      ctRule.removeFormula(0);
    }
    ctRule.setType(STCfType.TOP_10);
    ctRule.setRank(top10Rule.rank());
    ctRule.setPercent(top10Rule.percent());
    ctRule.setBottom(top10Rule.bottom());
  }

  private void clearAll(XSSFSheet sheet) {
    for (int blockIndex = sheet.getSheetConditionalFormatting().getNumConditionalFormattings() - 1;
        blockIndex >= 0;
        blockIndex--) {
      sheet.getSheetConditionalFormatting().removeConditionalFormatting(blockIndex);
    }
  }

  private void removeIntersectingBlocks(XSSFSheet sheet, List<ExcelRange> cutouts) {
    List<Integer> wrapperIndexesToRemove = new ArrayList<>();
    for (int blockIndex = 0;
        blockIndex < sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
        blockIndex++) {
      if (ExcelConditionalFormattingSnapshotSupport.intersectsAny(
          ExcelConditionalFormattingSnapshotSupport.formattingRanges(
              ExcelConditionalFormattingSnapshotSupport.ctBlock(sheet, blockIndex)),
          cutouts)) {
        wrapperIndexesToRemove.add(blockIndex);
      }
    }
    for (int index = wrapperIndexesToRemove.size() - 1; index >= 0; index--) {
      sheet
          .getSheetConditionalFormatting()
          .removeConditionalFormatting(wrapperIndexesToRemove.get(index));
    }
  }

  void normalizePriorities(XSSFSheet sheet) {
    int priority = 1;
    for (int blockIndex = 0;
        blockIndex < sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
        blockIndex++) {
      CTConditionalFormatting ctBlock =
          ExcelConditionalFormattingSnapshotSupport.ctBlock(sheet, blockIndex);
      for (int ruleIndex = 0; ruleIndex < ctBlock.sizeOfCfRuleArray(); ruleIndex++) {
        ctBlock.getCfRuleArray(ruleIndex).setPriority(priority);
        priority++;
      }
    }
  }
}
