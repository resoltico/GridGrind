package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import javax.xml.namespace.QName;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;

/** Reads conditional-formatting XML into canonical GridGrind snapshot values. */
final class ExcelConditionalFormattingSnapshotSupport {
  List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    List<ExcelConditionalFormattingBlockSnapshot> snapshots = new ArrayList<>();
    for (int blockIndex = 0;
        blockIndex < sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
        blockIndex++) {
      XSSFConditionalFormatting formatting =
          sheet.getSheetConditionalFormatting().getConditionalFormattingAt(blockIndex);
      CTConditionalFormatting ctBlock = ctBlock(sheet, blockIndex);
      List<String> ranges = formattingRanges(ctBlock);
      if (!matchesSelection(ranges, selection)) {
        continue;
      }
      snapshots.add(
          new ExcelConditionalFormattingBlockSnapshot(
              ranges, snapshotRules(sheet, formatting, ctBlock)));
    }
    return List.copyOf(snapshots);
  }

  static List<String> formattingRanges(CTConditionalFormatting ctBlock) {
    return ExcelSqrefSupport.normalizedSqref(ctBlock.getSqref());
  }

  static CTConditionalFormatting ctBlock(XSSFSheet sheet, int blockIndex) {
    return sheet.getCTWorksheet().getConditionalFormattingArray(blockIndex);
  }

  static boolean matchesSelection(List<String> ranges, ExcelRangeSelection selection) {
    return switch (selection) {
      case ExcelRangeSelection.All _ -> true;
      case ExcelRangeSelection.Selected selected ->
          intersectsAny(ranges, selected.ranges().stream().map(ExcelRange::parse).toList());
    };
  }

  static boolean intersectsAny(List<String> rawRanges, List<ExcelRange> targetRanges) {
    for (String rawRange : rawRanges) {
      ExcelRange existingRange = ExcelSheetStructureSupport.parseRangeOrNull(rawRange);
      if (existingRange == null) {
        continue;
      }
      for (ExcelRange targetRange : targetRanges) {
        if (ExcelSheetStructureSupport.intersects(existingRange, targetRange)) {
          return true;
        }
      }
    }
    return false;
  }

  static ExcelConditionalFormattingRuleSnapshot toSnapshot(
      XSSFSheet sheet, XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    ExcelDifferentialStyleSnapshot style =
        ExcelConditionalFormattingStyleSupport.snapshotStyle(
            sheet.getWorkbook().getStylesSource(), ctRule);

    String rawType = rawCfType(ctRule);
    if (rawType != null && "top10".equalsIgnoreCase(rawType)) {
      return top10RuleSnapshot(ctRule, style);
    }

    ConditionType conditionType;
    try {
      conditionType = rule.getConditionType();
    } catch (RuntimeException exception) {
      return unsupportedRule(ctRule, "UNKNOWN", "Rule family metadata is missing or unreadable.");
    }
    if (conditionType == null) {
      return unsupportedRule(ctRule, "UNKNOWN", "Rule family metadata is missing or unreadable.");
    }
    if (conditionType == ConditionType.FORMULA) {
      return formulaRuleSnapshot(rule, ctRule, style);
    }
    if (conditionType == ConditionType.CELL_VALUE_IS) {
      return cellValueRuleSnapshot(rule, ctRule, style);
    }
    if (conditionType == ConditionType.COLOR_SCALE) {
      return colorScaleRuleSnapshot(rule, ctRule);
    }
    if (conditionType == ConditionType.DATA_BAR) {
      return dataBarRuleSnapshot(rule, ctRule);
    }
    if (conditionType == ConditionType.ICON_SET) {
      return iconSetRuleSnapshot(rule, ctRule);
    }
    return unsupportedRule(
        ctRule, unsupportedKind(conditionType, ctRule), "Rule family is not modeled by GridGrind.");
  }

  static String unsupportedKind(ConditionType conditionType, CTCfRule ctRule) {
    String rawType = rawCfType(ctRule);
    if (rawType != null) {
      return normalizedUnsupportedKind(rawType);
    }
    if (conditionType == ConditionType.FORMULA) {
      return "FORMULA";
    }
    if (conditionType == ConditionType.CELL_VALUE_IS) {
      return "CELL_VALUE_IS";
    }
    if (conditionType == ConditionType.COLOR_SCALE) {
      return "COLOR_SCALE";
    }
    if (conditionType == ConditionType.DATA_BAR) {
      return "DATA_BAR";
    }
    if (conditionType == ConditionType.ICON_SET) {
      return "ICON_SET";
    }
    return normalizedUnsupportedKind(conditionType.toString());
  }

  static String normalizedUnsupportedKind(String rawKind) {
    return rawKind
        .replaceAll("([a-z])([A-Z])", "$1_$2")
        .replaceAll("([A-Za-z])(\\d)", "$1_$2")
        .toUpperCase(Locale.ROOT);
  }

  static boolean isBrokenFormula(
      XSSFEvaluationWorkbook evaluationWorkbook, int sheetIndex, String formula) {
    if (formula.contains("#REF!")) {
      return true;
    }
    try {
      org.apache.poi.ss.formula.FormulaParser.parse(
          formula, evaluationWorkbook, org.apache.poi.ss.formula.FormulaType.CELL, sheetIndex);
      return false;
    } catch (RuntimeException exception) {
      return true;
    }
  }

  static String locationEvidence(WorkbookAnalysis.AnalysisLocation location) {
    return switch (location) {
      case WorkbookAnalysis.AnalysisLocation.Workbook _ -> "WORKBOOK";
      case WorkbookAnalysis.AnalysisLocation.Sheet sheet -> sheet.sheetName();
      case WorkbookAnalysis.AnalysisLocation.Cell cell -> cell.sheetName() + "!" + cell.address();
      case WorkbookAnalysis.AnalysisLocation.Range range -> range.sheetName() + "!" + range.range();
      case WorkbookAnalysis.AnalysisLocation.NamedRange namedRange ->
          namedRange.name() + "@" + namedRange.scope();
    };
  }

  static String ruleLabel(ExcelConditionalFormattingRuleSnapshot rule) {
    return switch (rule) {
      case ExcelConditionalFormattingRuleSnapshot.FormulaRule _ -> "FORMULA_RULE";
      case ExcelConditionalFormattingRuleSnapshot.CellValueRule _ -> "CELL_VALUE_RULE";
      case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule _ -> "COLOR_SCALE_RULE";
      case ExcelConditionalFormattingRuleSnapshot.DataBarRule _ -> "DATA_BAR_RULE";
      case ExcelConditionalFormattingRuleSnapshot.IconSetRule _ -> "ICON_SET_RULE";
      case ExcelConditionalFormattingRuleSnapshot.Top10Rule _ -> "TOP10_RULE";
      case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
          "UNSUPPORTED_RULE(" + unsupportedRule.kind().toUpperCase(Locale.ROOT) + ")";
    };
  }

  private static List<ExcelConditionalFormattingRuleSnapshot> snapshotRules(
      XSSFSheet sheet, XSSFConditionalFormatting formatting, CTConditionalFormatting ctBlock) {
    List<ExcelConditionalFormattingRuleSnapshot> rules = new ArrayList<>();
    for (int ruleIndex = 0; ruleIndex < formatting.getNumberOfRules(); ruleIndex++) {
      rules.add(
          toSnapshot(sheet, formatting.getRule(ruleIndex), ctBlock.getCfRuleArray(ruleIndex)));
    }
    return List.copyOf(rules);
  }

  private static List<ExcelConditionalFormattingThresholdSnapshot> thresholds(
      ConditionalFormattingThreshold[] thresholds) {
    List<ExcelConditionalFormattingThresholdSnapshot> snapshots =
        new ArrayList<>(thresholds.length);
    for (ConditionalFormattingThreshold threshold : thresholds) {
      snapshots.add(threshold(threshold));
    }
    return List.copyOf(snapshots);
  }

  private static ExcelConditionalFormattingThresholdSnapshot threshold(
      ConditionalFormattingThreshold threshold) {
    return new ExcelConditionalFormattingThresholdSnapshot(
        ExcelConditionalFormattingPoiBridge.fromPoi(threshold.getRangeType()),
        threshold.getFormula(),
        threshold.getValue());
  }

  private static List<String> colors(XSSFColor[] colors) {
    List<String> values = new ArrayList<>(colors.length);
    for (XSSFColor color : colors) {
      values.add(color(color));
    }
    return List.copyOf(values);
  }

  private static String color(XSSFColor color) {
    String rgb = ExcelRgbColorSupport.toRgbHex(color).orElse(null);
    if (rgb == null) {
      throw new IllegalArgumentException("conditional-formatting color must expose RGB data");
    }
    return rgb;
  }

  private static ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule(
      CTCfRule ctRule, String kind, String detail) {
    return new ExcelConditionalFormattingRuleSnapshot.UnsupportedRule(
        ctRule.getPriority(), ctRule.getStopIfTrue(), kind, detail);
  }

  private static String rawCfType(CTCfRule ctRule) {
    try (var cursor = ctRule.newCursor()) {
      return cursor.getAttributeText(new QName("", "type"));
    }
  }

  private static ExcelConditionalFormattingRuleSnapshot formulaRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule, ExcelDifferentialStyleSnapshot style) {
    String formula = Objects.requireNonNullElse(rule.getFormula1(), "");
    if (formula.isBlank()) {
      return unsupportedRule(ctRule, "FORMULA", "Formula rule is missing formula text.");
    }
    return new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
        ctRule.getPriority(), ctRule.getStopIfTrue(), formula, style);
  }

  private static ExcelConditionalFormattingRuleSnapshot cellValueRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule, ExcelDifferentialStyleSnapshot style) {
    String formula1 = Objects.requireNonNullElse(rule.getFormula1(), "");
    if (formula1.isBlank()) {
      return unsupportedRule(ctRule, "CELL_VALUE_IS", "Cell-value rule is missing formula1.");
    }
    try {
      return new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          ExcelComparisonOperatorPoiBridge.fromPoi(rule.getComparisonOperation()),
          formula1,
          rule.getFormula2(),
          style);
    } catch (IllegalArgumentException exception) {
      return unsupportedRule(
          ctRule,
          "CELL_VALUE_IS",
          "Cell-value rule uses unsupported comparison payload: " + exception.getMessage());
    }
  }

  private static ExcelConditionalFormattingRuleSnapshot colorScaleRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    var colorScaleFormatting = rule.getColorScaleFormatting();
    if (colorScaleFormatting == null) {
      return unsupportedRule(
          ctRule, "COLOR_SCALE", "Color-scale rule is missing color-scale payload.");
    }
    try {
      return new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          thresholds(colorScaleFormatting.getThresholds()),
          colors(colorScaleFormatting.getColors()));
    } catch (RuntimeException exception) {
      return unsupportedRule(
          ctRule,
          "COLOR_SCALE",
          "Color-scale rule uses unsupported threshold or color payload: "
              + exception.getMessage());
    }
  }

  private static ExcelConditionalFormattingRuleSnapshot dataBarRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    var dataBarFormatting = rule.getDataBarFormatting();
    if (dataBarFormatting == null) {
      return unsupportedRule(ctRule, "DATA_BAR", "Data-bar rule is missing data-bar payload.");
    }
    try {
      return new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          color(dataBarFormatting.getColor()),
          dataBarFormatting.isIconOnly(),
          dataBarFormatting.getWidthMin(),
          dataBarFormatting.getWidthMax(),
          threshold(dataBarFormatting.getMinThreshold()),
          threshold(dataBarFormatting.getMaxThreshold()));
    } catch (RuntimeException exception) {
      return unsupportedRule(
          ctRule,
          "DATA_BAR",
          "Data-bar rule uses unsupported threshold or color payload: " + exception.getMessage());
    }
  }

  private static ExcelConditionalFormattingRuleSnapshot iconSetRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    var iconFormatting = rule.getMultiStateFormatting();
    if (iconFormatting == null) {
      return unsupportedRule(ctRule, "ICON_SET", "Icon-set rule is missing icon-set payload.");
    }
    try {
      return new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          ExcelConditionalFormattingPoiBridge.fromPoi(iconFormatting.getIconSet()),
          iconFormatting.isIconOnly(),
          iconFormatting.isReversed(),
          thresholds(iconFormatting.getThresholds()));
    } catch (RuntimeException exception) {
      return unsupportedRule(
          ctRule,
          "ICON_SET",
          "Icon-set rule uses unsupported threshold or icon-set payload: "
              + exception.getMessage());
    }
  }

  private static ExcelConditionalFormattingRuleSnapshot top10RuleSnapshot(
      CTCfRule ctRule, ExcelDifferentialStyleSnapshot style) {
    return new ExcelConditionalFormattingRuleSnapshot.Top10Rule(
        ctRule.getPriority(),
        ctRule.getStopIfTrue(),
        ctRule.isSetRank() ? Math.toIntExact(ctRule.getRank()) : 10,
        ctRule.isSetPercent() && ctRule.getPercent(),
        ctRule.isSetBottom() && ctRule.getBottom(),
        style);
  }
}
