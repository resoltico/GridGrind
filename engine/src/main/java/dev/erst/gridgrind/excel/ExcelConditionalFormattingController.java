package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import javax.xml.namespace.QName;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingThreshold;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;

/** Reads, writes, and analyzes conditional-formatting blocks on one XSSF sheet. */
final class ExcelConditionalFormattingController {
  /** Creates or replaces one logical conditional-formatting block on the target sheet. */
  void setConditionalFormatting(XSSFSheet sheet, ExcelConditionalFormattingBlockDefinition block) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(block, "block must not be null");

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

    CTConditionalFormatting ctBlock = ctBlock(sheet, blockIndex);
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

  /** Removes conditional-formatting blocks on one sheet according to the provided selection. */
  void clearConditionalFormatting(XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    switch (selection) {
      case ExcelRangeSelection.All _ -> clearAll(sheet);
      case ExcelRangeSelection.Selected selected ->
          removeIntersectingBlocks(
              sheet, selected.ranges().stream().map(ExcelRange::parse).toList());
    }
    normalizePriorities(sheet);
  }

  /** Returns factual conditional-formatting blocks for one sheet and range selection. */
  List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

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

  /** Returns the number of conditional-formatting blocks currently present on one sheet. */
  int conditionalFormattingBlockCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
  }

  /** Returns derived conditional-formatting health findings for one sheet. */
  List<WorkbookAnalysis.AnalysisFinding> conditionalFormattingHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");

    List<ExcelConditionalFormattingBlockSnapshot> blocks =
        conditionalFormatting(sheet, new ExcelRangeSelection.All());
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    List<PriorityContext> priorities = new ArrayList<>();
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(sheet.getWorkbook());

    for (ExcelConditionalFormattingBlockSnapshot block : blocks) {
      if (block.ranges().isEmpty() || hasInvalidRanges(block.ranges())) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
                WorkbookAnalysis.AnalysisSeverity.WARNING,
                "Conditional-formatting block targets an empty or invalid range",
                "Conditional-formatting block has no valid target ranges.",
                blockLocation(sheetName, block.ranges()),
                List.copyOf(block.ranges())));
      }

      for (ExcelConditionalFormattingRuleSnapshot rule : block.rules()) {
        priorities.add(priorityContext(sheetName, block, rule));
        switch (rule) {
          case ExcelConditionalFormattingRuleSnapshot.FormulaRule formulaRule -> {
            if (isBrokenFormula(evaluationWorkbook, sheetIndex, formulaRule.formula())) {
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
                      WorkbookAnalysis.AnalysisSeverity.ERROR,
                      "Conditional-formatting formula is invalid",
                      "Formula rule could not be parsed for conditional-formatting evaluation.",
                      blockLocation(sheetName, block.ranges()),
                      List.of(formulaRule.formula())));
            }
          }
          case ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule -> {
            if (isBrokenFormula(evaluationWorkbook, sheetIndex, cellValueRule.formula1())
                || (cellValueRule.formula2() != null
                    && isBrokenFormula(evaluationWorkbook, sheetIndex, cellValueRule.formula2()))) {
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
                      WorkbookAnalysis.AnalysisSeverity.ERROR,
                      "Conditional-formatting operand formula is invalid",
                      "Cell-value rule operands could not be parsed for conditional-formatting evaluation.",
                      blockLocation(sheetName, block.ranges()),
                      cellValueRuleEvidence(cellValueRule)));
            }
          }
          case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_UNSUPPORTED_RULE,
                      WorkbookAnalysis.AnalysisSeverity.WARNING,
                      "Unsupported conditional-formatting rule",
                      unsupportedRule.detail(),
                      blockLocation(sheetName, block.ranges()),
                      List.of(unsupportedRule.kind())));
          case ExcelConditionalFormattingRuleSnapshot.ColorScaleRule _ -> {}
          case ExcelConditionalFormattingRuleSnapshot.DataBarRule _ -> {}
          case ExcelConditionalFormattingRuleSnapshot.IconSetRule _ -> {}
          case ExcelConditionalFormattingRuleSnapshot.Top10Rule _ -> {}
        }
      }
    }

    findings.addAll(priorityCollisionFindings(priorities));

    return List.copyOf(findings);
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
      if (intersectsAny(formattingRanges(ctBlock(sheet, blockIndex)), cutouts)) {
        wrapperIndexesToRemove.add(blockIndex);
      }
    }
    for (int index = wrapperIndexesToRemove.size() - 1; index >= 0; index--) {
      sheet
          .getSheetConditionalFormatting()
          .removeConditionalFormatting(wrapperIndexesToRemove.get(index));
    }
  }

  private void normalizePriorities(XSSFSheet sheet) {
    int priority = 1;
    for (int blockIndex = 0;
        blockIndex < sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
        blockIndex++) {
      CTConditionalFormatting ctBlock = ctBlock(sheet, blockIndex);
      for (int ruleIndex = 0; ruleIndex < ctBlock.sizeOfCfRuleArray(); ruleIndex++) {
        ctBlock.getCfRuleArray(ruleIndex).setPriority(priority);
        priority++;
      }
    }
  }

  private static List<String> formattingRanges(CTConditionalFormatting ctBlock) {
    return ExcelSqrefSupport.normalizedSqref(ctBlock.getSqref());
  }

  private static CTConditionalFormatting ctBlock(XSSFSheet sheet, int blockIndex) {
    return sheet.getCTWorksheet().getConditionalFormattingArray(blockIndex);
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

  private boolean matchesSelection(List<String> ranges, ExcelRangeSelection selection) {
    return switch (selection) {
      case ExcelRangeSelection.All _ -> true;
      case ExcelRangeSelection.Selected selected ->
          intersectsAny(ranges, selected.ranges().stream().map(ExcelRange::parse).toList());
    };
  }

  private boolean intersectsAny(List<String> rawRanges, List<ExcelRange> targetRanges) {
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

  private static boolean hasInvalidRanges(List<String> ranges) {
    for (String range : ranges) {
      if (ExcelSheetStructureSupport.parseRangeOrNull(range) == null) {
        return true;
      }
    }
    return false;
  }

  /** Converts one loaded POI conditional-formatting rule into the matching GridGrind snapshot. */
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

  private static PriorityContext priorityContext(
      String sheetName,
      ExcelConditionalFormattingBlockSnapshot block,
      ExcelConditionalFormattingRuleSnapshot rule) {
    return new PriorityContext(
        rule.priority(),
        new BlockRuleContext(
            blockLocation(sheetName, block.ranges()), block.ranges(), ruleLabel(rule)));
  }

  private static List<String> cellValueRuleEvidence(
      ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule) {
    if (cellValueRule.formula2() == null) {
      return List.of(cellValueRule.formula1());
    }
    return List.of(cellValueRule.formula1(), cellValueRule.formula2());
  }

  private static List<WorkbookAnalysis.AnalysisFinding> priorityCollisionFindings(
      List<PriorityContext> priorities) {
    List<PriorityContext> sorted = new ArrayList<>(priorities);
    sorted.sort(java.util.Comparator.comparingInt(PriorityContext::priority));
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();

    int index = 0;
    while (index < sorted.size()) {
      int priority = sorted.get(index).priority();
      int nextIndex = index + 1;
      while (nextIndex < sorted.size() && sorted.get(nextIndex).priority() == priority) {
        nextIndex++;
      }
      if (priority <= 0 || nextIndex - index > 1) {
        findings.add(priorityCollisionFinding(sorted.subList(index, nextIndex)));
      }
      index = nextIndex;
    }

    return List.copyOf(findings);
  }

  private static WorkbookAnalysis.AnalysisFinding priorityCollisionFinding(
      List<PriorityContext> contexts) {
    List<String> evidence =
        contexts.stream()
            .map(PriorityContext::context)
            .map(context -> context.label() + "@" + locationEvidence(context.location()))
            .toList();
    return new WorkbookAnalysis.AnalysisFinding(
        WorkbookAnalysis.AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
        WorkbookAnalysis.AnalysisSeverity.WARNING,
        "Conditional-formatting priorities collide",
        "Conditional-formatting priorities must be positive and unique within a sheet.",
        contexts.getFirst().context().location(),
        evidence);
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

  /** Returns the stable unsupported-rule kind label for one unmodeled conditional-format family. */
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

  private static String rawCfType(CTCfRule ctRule) {
    try (var cursor = ctRule.newCursor()) {
      return cursor.getAttributeText(new QName("", "type"));
    }
  }

  /** Normalizes one raw conditional-format family name into GridGrind's stable uppercase kind. */
  static String normalizedUnsupportedKind(String rawKind) {
    return rawKind
        .replaceAll("([a-z])([A-Z])", "$1_$2")
        .replaceAll("([A-Za-z])(\\d)", "$1_$2")
        .toUpperCase(Locale.ROOT);
  }

  /** Reports whether one conditional-formatting formula is syntactically broken for evaluation. */
  static boolean isBrokenFormula(
      XSSFEvaluationWorkbook evaluationWorkbook, int sheetIndex, String formula) {
    if (formula.contains("#REF!")) {
      return true;
    }
    try {
      FormulaParser.parse(formula, evaluationWorkbook, FormulaType.CELL, sheetIndex);
      return false;
    } catch (RuntimeException exception) {
      return true;
    }
  }

  private static WorkbookAnalysis.AnalysisLocation blockLocation(
      String sheetName, List<String> ranges) {
    if (!ranges.isEmpty()) {
      return new WorkbookAnalysis.AnalysisLocation.Range(sheetName, ranges.getFirst());
    }
    return new WorkbookAnalysis.AnalysisLocation.Sheet(sheetName);
  }

  /** Formats one analysis location into the evidence string used by collision findings. */
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

  /** Returns the stable label used in priority-collision evidence for one snapshot rule family. */
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

  private record BlockRuleContext(
      WorkbookAnalysis.AnalysisLocation location, List<String> ranges, String label) {}

  private record PriorityContext(int priority, BlockRuleContext context) {}
}
