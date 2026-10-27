package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import javax.xml.namespace.QName;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;

/** Reads conditional-formatting XML into canonical GridGrind snapshot values. */
final class ExcelConditionalFormattingSnapshotSupport {
  List<ConditionalFormattingEntryReport> conditionalFormatting(XSSFSheet sheet) {
    return readConditionalFormatting(sheet, ranges -> true);
  }

  List<ConditionalFormattingEntryReport> conditionalFormatting(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> targets) {
    var copied = List.copyOf(targets);
    return readConditionalFormatting(sheet, ranges -> matchesSelection(ranges, copied));
  }

  private List<ConditionalFormattingEntryReport> readConditionalFormatting(
      XSSFSheet sheet, java.util.function.Predicate<List<String>> matches) {
    List<ConditionalFormattingEntryReport> snapshots = new ArrayList<>();
    for (int blockIndex = 0;
        blockIndex < sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
        blockIndex++) {
      XSSFConditionalFormatting formatting =
          sheet.getSheetConditionalFormatting().getConditionalFormattingAt(blockIndex);
      CTConditionalFormatting ctBlock = ctBlock(sheet, blockIndex);
      List<String> ranges = formattingRanges(ctBlock);
      if (!matches.test(ranges)) {
        continue;
      }
      snapshots.add(
          new ConditionalFormattingEntryReport(
              ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
              snapshotRules(sheet, formatting, ctBlock)));
    }
    return List.copyOf(snapshots);
  }

  static List<String> formattingRanges(CTConditionalFormatting ctBlock) {
    return ExcelSqrefSupport.normalizedSqref(ctBlock.getSqref());
  }

  static CTConditionalFormatting ctBlock(XSSFSheet sheet, int blockIndex) {
    return sheet.getCTWorksheet().getConditionalFormattingArray(blockIndex);
  }

  static boolean matchesSelection(
      List<String> ranges, List<dev.erst.gridgrind.core.vocabulary.RangeRef> targets) {
    return intersectsAny(
        ranges, targets.stream().map(target -> ExcelRange.parse(target.value())).toList());
  }

  static boolean intersectsAny(List<String> rawRanges, List<ExcelRange> targetRanges) {
    for (String rawRange : rawRanges) {
      Optional<ExcelRange> existingRange = ExcelSheetStructureSupport.parseOptionalRange(rawRange);
      if (existingRange.isEmpty()) {
        continue;
      }
      for (ExcelRange targetRange : targetRanges) {
        if (ExcelSheetStructureSupport.intersects(existingRange.orElseThrow(), targetRange)) {
          return true;
        }
      }
    }
    return false;
  }

  static ConditionalFormattingRuleReport toSnapshot(
      XSSFSheet sheet, XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    DifferentialStyleReport style =
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
    if (ConditionType.FORMULA.equals(conditionType)) {
      return formulaRuleSnapshot(rule, ctRule, style);
    }
    if (ConditionType.CELL_VALUE_IS.equals(conditionType)) {
      return cellValueRuleSnapshot(rule, ctRule, style);
    }
    if (ConditionType.COLOR_SCALE.equals(conditionType)) {
      return colorScaleRuleSnapshot(rule, ctRule);
    }
    if (ConditionType.DATA_BAR.equals(conditionType)) {
      return dataBarRuleSnapshot(rule, ctRule);
    }
    if (ConditionType.ICON_SET.equals(conditionType)) {
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
    if (ConditionType.FORMULA.equals(conditionType)) {
      return "FORMULA";
    }
    if (ConditionType.CELL_VALUE_IS.equals(conditionType)) {
      return "CELL_VALUE_IS";
    }
    if (ConditionType.COLOR_SCALE.equals(conditionType)) {
      return "COLOR_SCALE";
    }
    if (ConditionType.DATA_BAR.equals(conditionType)) {
      return "DATA_BAR";
    }
    if (ConditionType.ICON_SET.equals(conditionType)) {
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

  static String locationEvidence(AnalysisLocationReport location) {
    return switch (location) {
      case AnalysisLocationReport.Workbook _ -> "WORKBOOK";
      case AnalysisLocationReport.Sheet sheet -> sheet.sheetName().value();
      case AnalysisLocationReport.Cell cell -> cell.sheetName() + "!" + cell.address();
      case AnalysisLocationReport.Range range -> range.sheetName() + "!" + range.range();
      case AnalysisLocationReport.NamedRange namedRange ->
          namedRange.name()
              + "@"
              + switch (namedRange.scope()) {
                case dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Workbook _ ->
                    "workbook";
                case dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Sheet sheet ->
                    "sheet:" + sheet.sheetName().value();
              };
    };
  }

  static String ruleLabel(ConditionalFormattingRuleReport rule) {
    return switch (rule) {
      case ConditionalFormattingRuleReport.FormulaRule _ -> "FORMULA_RULE";
      case ConditionalFormattingRuleReport.CellValueRule _ -> "CELL_VALUE_RULE";
      case ConditionalFormattingRuleReport.ColorScaleRule _ -> "COLOR_SCALE_RULE";
      case ConditionalFormattingRuleReport.DataBarRule _ -> "DATA_BAR_RULE";
      case ConditionalFormattingRuleReport.IconSetRule _ -> "ICON_SET_RULE";
      case ConditionalFormattingRuleReport.Top10Rule _ -> "TOP10_RULE";
      case ConditionalFormattingRuleReport.UnsupportedRule unsupportedRule ->
          "UNSUPPORTED_RULE(" + unsupportedRule.kind().toUpperCase(Locale.ROOT) + ")";
    };
  }

  private static List<ConditionalFormattingRuleReport> snapshotRules(
      XSSFSheet sheet, XSSFConditionalFormatting formatting, CTConditionalFormatting ctBlock) {
    List<ConditionalFormattingRuleReport> rules = new ArrayList<>();
    for (int ruleIndex = 0; ruleIndex < formatting.getNumberOfRules(); ruleIndex++) {
      rules.add(
          toSnapshot(sheet, formatting.getRule(ruleIndex), ctBlock.getCfRuleArray(ruleIndex)));
    }
    return List.copyOf(rules);
  }

  private static List<ConditionalFormattingThresholdReport> thresholds(
      ConditionalFormattingThreshold[] thresholds) {
    List<ConditionalFormattingThresholdReport> snapshots = new ArrayList<>(thresholds.length);
    for (ConditionalFormattingThreshold threshold : thresholds) {
      snapshots.add(threshold(threshold));
    }
    return List.copyOf(snapshots);
  }

  private static ConditionalFormattingThresholdReport threshold(
      ConditionalFormattingThreshold threshold) {
    return new ConditionalFormattingThresholdReport(
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

  private static ConditionalFormattingRuleReport.UnsupportedRule unsupportedRule(
      CTCfRule ctRule, String kind, String detail) {
    return new ConditionalFormattingRuleReport.UnsupportedRule(
        ctRule.getPriority(), ctRule.getStopIfTrue(), kind, detail);
  }

  private static String rawCfType(CTCfRule ctRule) {
    try (var cursor = ctRule.newCursor()) {
      return cursor.getAttributeText(new QName("", "type"));
    }
  }

  private static ConditionalFormattingRuleReport formulaRuleSnapshot(
      XSSFConditionalFormattingRule rule,
      CTCfRule ctRule,
      @Nullable DifferentialStyleReport style) {
    String formula = Objects.requireNonNullElse(rule.getFormula1(), "");
    if (formula.isBlank()) {
      return unsupportedRule(ctRule, "FORMULA", "Formula rule is missing formula text.");
    }
    return new ConditionalFormattingRuleReport.FormulaRule(
        ctRule.getPriority(), ctRule.getStopIfTrue(), formula, Optional.ofNullable(style));
  }

  private static ConditionalFormattingRuleReport cellValueRuleSnapshot(
      XSSFConditionalFormattingRule rule,
      CTCfRule ctRule,
      @Nullable DifferentialStyleReport style) {
    String formula1 = Objects.requireNonNullElse(rule.getFormula1(), "");
    if (formula1.isBlank()) {
      return unsupportedRule(ctRule, "CELL_VALUE_IS", "Cell-value rule is missing formula1.");
    }
    try {
      return new ConditionalFormattingRuleReport.CellValueRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(
              rule.getComparisonOperation()),
          formula1,
          Optional.ofNullable(rule.getFormula2()),
          Optional.ofNullable(style));
    } catch (IllegalArgumentException exception) {
      return unsupportedRule(
          ctRule,
          "CELL_VALUE_IS",
          "Cell-value rule uses unsupported comparison payload: " + exception.getMessage());
    }
  }

  private static ConditionalFormattingRuleReport colorScaleRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    var colorScaleFormatting = rule.getColorScaleFormatting();
    if (colorScaleFormatting == null) {
      return unsupportedRule(
          ctRule, "COLOR_SCALE", "Color-scale rule is missing color-scale payload.");
    }
    try {
      return new ConditionalFormattingRuleReport.ColorScaleRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          thresholds(colorScaleFormatting.getThresholds()),
          colors(colorScaleFormatting.getColors()).stream()
              .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new)
              .toList());
    } catch (RuntimeException exception) {
      return unsupportedRule(
          ctRule,
          "COLOR_SCALE",
          "Color-scale rule uses unsupported threshold or color payload: "
              + exception.getMessage());
    }
  }

  private static ConditionalFormattingRuleReport dataBarRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    var dataBarFormatting = rule.getDataBarFormatting();
    if (dataBarFormatting == null) {
      return unsupportedRule(ctRule, "DATA_BAR", "Data-bar rule is missing data-bar payload.");
    }
    try {
      return new ConditionalFormattingRuleReport.DataBarRule(
          ctRule.getPriority(),
          ctRule.getStopIfTrue(),
          Optional.ofNullable(color(dataBarFormatting.getColor()))
              .map(dev.erst.gridgrind.core.vocabulary.RgbHex::new),
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

  private static ConditionalFormattingRuleReport iconSetRuleSnapshot(
      XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    var iconFormatting = rule.getMultiStateFormatting();
    if (iconFormatting == null) {
      return unsupportedRule(ctRule, "ICON_SET", "Icon-set rule is missing icon-set payload.");
    }
    try {
      return new ConditionalFormattingRuleReport.IconSetRule(
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

  private static ConditionalFormattingRuleReport top10RuleSnapshot(
      CTCfRule ctRule, @Nullable DifferentialStyleReport style) {
    return new ConditionalFormattingRuleReport.Top10Rule(
        ctRule.getPriority(),
        ctRule.getStopIfTrue(),
        ctRule.isSetRank() ? Math.toIntExact(ctRule.getRank()) : 10,
        ctRule.isSetPercent() && ctRule.getPercent(),
        ctRule.isSetBottom() && ctRule.getBottom(),
        Optional.ofNullable(style));
  }
}
