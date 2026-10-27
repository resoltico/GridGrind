package dev.erst.gridgrind.engine.conditionalformatting;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderSideReport;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.EnumMappings;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.poi.SqrefRanges;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import javax.xml.namespace.QName;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingThreshold;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontSize;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

/** Owns conditional-formatting rules, differential styles, raw-state readback and diagnostics. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class ConditionalFormattingAdapter {
  public int conditionalFormattingBlockCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
  }

  public List<AnalysisFindingReport> conditionalFormattingHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    return analyzeBlocks(sheetName, sheet, conditionalFormatting(sheet));
  }

  public void setConditionalFormatting(
      XSSFSheet sheet,
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges,
      dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec formatting) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(formatting, "formatting must not be null");
    var targets = List.copyOf(ranges);
    if (targets.isEmpty()) throw new IllegalArgumentException("ranges must not be empty");
    if (new java.util.LinkedHashSet<>(targets).size() != targets.size())
      throw new IllegalArgumentException("ranges must not contain duplicates");
    List<CellBounds> targetRanges =
        targets.stream().map(range -> CellBounds.parse(range.value())).toList();
    XSSFConditionalFormattingRule[] rules =
        formatting.rules().stream()
            .map(rule -> createRule(sheet, rule))
            .toArray(XSSFConditionalFormattingRule[]::new);
    removeIntersectingBlocks(sheet, targetRanges);
    int blockIndex =
        sheet
            .getSheetConditionalFormatting()
            .addConditionalFormatting(
                targetRanges.stream()
                    .map(SheetStructure::toCellRangeAddress)
                    .toArray(CellRangeAddress[]::new),
                rules);

    CTConditionalFormatting ctBlock = ctBlock(sheet, blockIndex);
    for (int ruleIndex = 0; ruleIndex < formatting.rules().size(); ruleIndex++) {
      CTCfRule ctRule = ctBlock.getCfRuleArray(ruleIndex);
      switch (formatting.rules().get(ruleIndex)) {
        case ConditionalFormattingRuleSpec.FormulaRule formulaRule -> {
          ctRule.setStopIfTrue(formulaRule.stopIfTrue());
          formulaRule.style().ifPresent(style -> applyStyle(sheet.getWorkbook(), ctRule, style));
        }
        case ConditionalFormattingRuleSpec.CellValueRule cellValueRule -> {
          ctRule.setStopIfTrue(cellValueRule.stopIfTrue());
          cellValueRule.style().ifPresent(style -> applyStyle(sheet.getWorkbook(), ctRule, style));
        }
        case ConditionalFormattingRuleSpec.ColorScaleRule colorScaleRule ->
            ctRule.setStopIfTrue(colorScaleRule.stopIfTrue());
        case ConditionalFormattingRuleSpec.DataBarRule dataBarRule ->
            ctRule.setStopIfTrue(dataBarRule.stopIfTrue());
        case ConditionalFormattingRuleSpec.IconSetRule iconSetRule ->
            ctRule.setStopIfTrue(iconSetRule.stopIfTrue());
        case ConditionalFormattingRuleSpec.Top10Rule top10Rule -> {
          ctRule.setStopIfTrue(top10Rule.stopIfTrue());
          top10Rule.style().ifPresent(style -> applyStyle(sheet.getWorkbook(), ctRule, style));
          applyTop10Rule(ctRule, top10Rule);
        }
      }
    }
    normalizePriorities(sheet);
  }

  public void clearConditionalFormatting(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    clearAll(sheet);
    normalizePriorities(sheet);
  }

  public void clearConditionalFormatting(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    removeIntersectingBlocks(
        sheet, ranges.stream().map(range -> CellBounds.parse(range.value())).toList());
    normalizePriorities(sheet);
  }

  private XSSFConditionalFormattingRule createRule(
      XSSFSheet sheet, ConditionalFormattingRuleSpec rule) {
    return switch (rule) {
      case ConditionalFormattingRuleSpec.FormulaRule formulaRule ->
          sheet
              .getSheetConditionalFormatting()
              .createConditionalFormattingRule(
                  ComparisonFormulas.normalizeFormula(formulaRule.formula(), "formula"));
      case ConditionalFormattingRuleSpec.CellValueRule cellValueRule ->
          cellValueRule.formula2().isEmpty()
              ? sheet
                  .getSheetConditionalFormatting()
                  .createConditionalFormattingRule(
                      toPoi(cellValueRule.operator()),
                      ComparisonFormulas.normalizeFormula(cellValueRule.formula1(), "formula1"))
              : sheet
                  .getSheetConditionalFormatting()
                  .createConditionalFormattingRule(
                      toPoi(cellValueRule.operator()),
                      ComparisonFormulas.normalizeFormula(cellValueRule.formula1(), "formula1"),
                      ComparisonFormulas.normalizeFormula(
                          cellValueRule.formula2().orElseThrow(), "formula2"));
      case ConditionalFormattingRuleSpec.ColorScaleRule colorScaleRule ->
          createColorScaleRule(sheet, colorScaleRule);
      case ConditionalFormattingRuleSpec.DataBarRule dataBarRule ->
          createDataBarRule(sheet, dataBarRule);
      case ConditionalFormattingRuleSpec.IconSetRule iconSetRule ->
          createIconSetRule(sheet, iconSetRule);
      case ConditionalFormattingRuleSpec.Top10Rule _ ->
          sheet
              .getSheetConditionalFormatting()
              .createConditionalFormattingRule(ComparisonOperator.GT, "0");
    };
  }

  private XSSFConditionalFormattingRule createColorScaleRule(
      XSSFSheet sheet, ConditionalFormattingRuleSpec.ColorScaleRule rule) {
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
            .map(color -> ColorAdapter.toXssfColor(sheet.getWorkbook(), color))
            .toArray(XSSFColor[]::new));
    return created;
  }

  private XSSFConditionalFormattingRule createDataBarRule(
      XSSFSheet sheet, ConditionalFormattingRuleSpec.DataBarRule rule) {
    XSSFConditionalFormattingRule created =
        sheet
            .getSheetConditionalFormatting()
            .createConditionalFormattingRule(
                ColorAdapter.toXssfColor(sheet.getWorkbook(), rule.color()));
    var formatting = created.getDataBarFormatting();
    formatting.setIconOnly(rule.iconOnly());
    formatting.setWidthMin(rule.widthMin());
    formatting.setWidthMax(rule.widthMax());
    applyThreshold(formatting.getMinThreshold(), rule.minThreshold());
    applyThreshold(formatting.getMaxThreshold(), rule.maxThreshold());
    return created;
  }

  private XSSFConditionalFormattingRule createIconSetRule(
      XSSFSheet sheet, ConditionalFormattingRuleSpec.IconSetRule rule) {
    XSSFConditionalFormattingRule created =
        sheet
            .getSheetConditionalFormatting()
            .createConditionalFormattingRule(toPoi(rule.iconSet()));
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
      ConditionalFormattingThresholdSpec authoredThreshold) {
    switch (authoredThreshold) {
      case ConditionalFormattingThresholdSpec.Min _ ->
          threshold.setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
      case ConditionalFormattingThresholdSpec.Max _ ->
          threshold.setRangeType(ConditionalFormattingThreshold.RangeType.MAX);
      case ConditionalFormattingThresholdSpec.Numeric value -> {
        threshold.setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
        threshold.setValue(value.value());
      }
      case ConditionalFormattingThresholdSpec.Percent value -> {
        threshold.setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
        threshold.setValue(value.value());
      }
      case ConditionalFormattingThresholdSpec.Percentile value -> {
        threshold.setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
        threshold.setValue(value.value());
      }
      case ConditionalFormattingThresholdSpec.Formula value -> {
        threshold.setRangeType(ConditionalFormattingThreshold.RangeType.FORMULA);
        threshold.setFormula(value.formula());
      }
    }
  }

  private static void applyTop10Rule(
      CTCfRule ctRule, ConditionalFormattingRuleSpec.Top10Rule top10Rule) {
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

  private void removeIntersectingBlocks(XSSFSheet sheet, List<CellBounds> cutouts) {
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

  private List<AnalysisFindingReport> analyzeBlocks(
      String sheetName, XSSFSheet sheet, List<ConditionalFormattingEntryReport> blocks) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    List<PriorityContext> priorities = new ArrayList<>();
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(sheet.getWorkbook());

    for (ConditionalFormattingEntryReport block : blocks) {
      addBlockRangeFindings(findings, sheetName, block);
      for (ConditionalFormattingRuleReport rule : block.rules()) {
        priorities.add(priorityContext(sheetName, block, rule));
        addRuleFindings(findings, evaluationWorkbook, sheetIndex, sheetName, block, rule);
      }
    }

    findings.addAll(priorityCollisionFindings(priorities));
    return List.copyOf(findings);
  }

  private static void addBlockRangeFindings(
      List<AnalysisFindingReport> findings,
      String sheetName,
      ConditionalFormattingEntryReport block) {
    if (block.ranges().isEmpty() || hasInvalidRanges(block.ranges())) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
              AnalysisSeverity.WARNING,
              "Conditional-formatting block targets an empty or invalid range",
              "Conditional-formatting block has no valid target ranges.",
              blockLocation(sheetName, block.ranges()),
              block.ranges().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
                  .toList()));
    }
  }

  private static void addRuleFindings(
      List<AnalysisFindingReport> findings,
      XSSFEvaluationWorkbook evaluationWorkbook,
      int sheetIndex,
      String sheetName,
      ConditionalFormattingEntryReport block,
      ConditionalFormattingRuleReport rule) {
    switch (rule) {
      case ConditionalFormattingRuleReport.FormulaRule formulaRule ->
          addFormulaRuleFinding(
              findings, evaluationWorkbook, sheetIndex, sheetName, block, formulaRule);
      case ConditionalFormattingRuleReport.CellValueRule cellValueRule ->
          addCellValueRuleFinding(
              findings, evaluationWorkbook, sheetIndex, sheetName, block, cellValueRule);
      case ConditionalFormattingRuleReport.UnsupportedRule unsupportedRule ->
          findings.add(
              new AnalysisFindingReport(
                  AnalysisFindingCode.CONDITIONAL_FORMATTING_UNSUPPORTED_RULE,
                  AnalysisSeverity.WARNING,
                  "Unsupported conditional-formatting rule",
                  unsupportedRule.detail(),
                  blockLocation(sheetName, block.ranges()),
                  List.of(unsupportedRule.kind())));
      case ConditionalFormattingRuleReport.ColorScaleRule _ -> {}
      case ConditionalFormattingRuleReport.DataBarRule _ -> {}
      case ConditionalFormattingRuleReport.IconSetRule _ -> {}
      case ConditionalFormattingRuleReport.Top10Rule _ -> {}
    }
  }

  private static void addFormulaRuleFinding(
      List<AnalysisFindingReport> findings,
      XSSFEvaluationWorkbook evaluationWorkbook,
      int sheetIndex,
      String sheetName,
      ConditionalFormattingEntryReport block,
      ConditionalFormattingRuleReport.FormulaRule formulaRule) {
    if (!isBrokenFormula(evaluationWorkbook, sheetIndex, formulaRule.formula())) {
      return;
    }
    findings.add(
        new AnalysisFindingReport(
            AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
            AnalysisSeverity.ERROR,
            "Conditional-formatting formula is invalid",
            "Formula rule could not be parsed for conditional-formatting evaluation.",
            blockLocation(sheetName, block.ranges()),
            List.of(formulaRule.formula())));
  }

  private static void addCellValueRuleFinding(
      List<AnalysisFindingReport> findings,
      XSSFEvaluationWorkbook evaluationWorkbook,
      int sheetIndex,
      String sheetName,
      ConditionalFormattingEntryReport block,
      ConditionalFormattingRuleReport.CellValueRule cellValueRule) {
    if (!hasBrokenCellValueOperand(evaluationWorkbook, sheetIndex, cellValueRule)) {
      return;
    }
    findings.add(
        new AnalysisFindingReport(
            AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
            AnalysisSeverity.ERROR,
            "Conditional-formatting operand formula is invalid",
            "Cell-value rule operands could not be parsed for conditional-formatting evaluation.",
            blockLocation(sheetName, block.ranges()),
            cellValueRuleEvidence(cellValueRule)));
  }

  private static boolean hasBrokenCellValueOperand(
      XSSFEvaluationWorkbook evaluationWorkbook,
      int sheetIndex,
      ConditionalFormattingRuleReport.CellValueRule cellValueRule) {
    return isBrokenFormula(evaluationWorkbook, sheetIndex, cellValueRule.formula1())
        || (cellValueRule.formula2().isPresent()
            && isBrokenFormula(
                evaluationWorkbook, sheetIndex, cellValueRule.formula2().orElseThrow()));
  }

  private static boolean hasInvalidRanges(
      List<dev.erst.gridgrind.core.vocabulary.RangeLabel> ranges) {
    for (var range : ranges) {
      if (SheetStructure.parseOptionalRange(range.value()).isEmpty()) {
        return true;
      }
    }
    return false;
  }

  private static PriorityContext priorityContext(
      String sheetName,
      ConditionalFormattingEntryReport block,
      ConditionalFormattingRuleReport rule) {
    return new PriorityContext(
        rule.priority(),
        new BlockRuleContext(
            blockLocation(sheetName, block.ranges()), block.ranges(), ruleLabel(rule)));
  }

  private static List<String> cellValueRuleEvidence(
      ConditionalFormattingRuleReport.CellValueRule cellValueRule) {
    if (cellValueRule.formula2().isEmpty()) {
      return List.of(cellValueRule.formula1());
    }
    return List.of(cellValueRule.formula1(), cellValueRule.formula2().orElseThrow());
  }

  private static List<AnalysisFindingReport> priorityCollisionFindings(
      List<PriorityContext> priorities) {
    List<PriorityContext> sorted = new ArrayList<>(priorities);
    sorted.sort(java.util.Comparator.comparingInt(PriorityContext::priority));
    List<AnalysisFindingReport> findings = new ArrayList<>();

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

  private static AnalysisFindingReport priorityCollisionFinding(List<PriorityContext> contexts) {
    List<String> evidence =
        contexts.stream()
            .map(PriorityContext::context)
            .map(context -> context.label() + "@" + locationEvidence(context.location()))
            .toList();
    return new AnalysisFindingReport(
        AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
        AnalysisSeverity.WARNING,
        "Conditional-formatting priorities collide",
        "Conditional-formatting priorities must be positive and unique within a sheet.",
        contexts.getFirst().context().location(),
        evidence);
  }

  private static AnalysisLocationReport blockLocation(
      String sheetName, List<dev.erst.gridgrind.core.vocabulary.RangeLabel> ranges) {
    if (!ranges.isEmpty()) {
      return new AnalysisLocationReport.Range(
          new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), ranges.getFirst().value());
    }
    return new AnalysisLocationReport.Sheet(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName));
  }

  private record BlockRuleContext(
      AnalysisLocationReport location,
      List<dev.erst.gridgrind.core.vocabulary.RangeLabel> ranges,
      String label) {
    private BlockRuleContext {
      ranges = List.copyOf(ranges);
    }
  }

  private record PriorityContext(int priority, BlockRuleContext context) {}

  public List<ConditionalFormattingEntryReport> conditionalFormatting(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return readConditionalFormatting(sheet, ranges -> true);
  }

  public List<ConditionalFormattingEntryReport> conditionalFormatting(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> targets) {
    Objects.requireNonNull(sheet, "sheet must not be null");
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

  private static List<String> formattingRanges(CTConditionalFormatting ctBlock) {
    return SqrefRanges.normalizedSqref(ctBlock.getSqref());
  }

  public static CTConditionalFormatting ctBlock(XSSFSheet sheet, int blockIndex) {
    return sheet.getCTWorksheet().getConditionalFormattingArray(blockIndex);
  }

  public static boolean matchesSelection(
      List<String> ranges, List<dev.erst.gridgrind.core.vocabulary.RangeRef> targets) {
    return intersectsAny(
        ranges, targets.stream().map(target -> CellBounds.parse(target.value())).toList());
  }

  private static boolean intersectsAny(List<String> rawRanges, List<CellBounds> targetRanges) {
    for (String rawRange : rawRanges) {
      Optional<CellBounds> existingRange = SheetStructure.parseOptionalRange(rawRange);
      if (existingRange.isEmpty()) {
        continue;
      }
      for (CellBounds targetRange : targetRanges) {
        if (SheetStructure.intersects(existingRange.orElseThrow(), targetRange)) {
          return true;
        }
      }
    }
    return false;
  }

  public static ConditionalFormattingRuleReport toSnapshot(
      XSSFSheet sheet, XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    DifferentialStyleReport style = snapshotStyle(sheet.getWorkbook().getStylesSource(), ctRule);

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

  public static String unsupportedKind(ConditionType conditionType, CTCfRule ctRule) {
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

  public static String normalizedUnsupportedKind(String rawKind) {
    return rawKind
        .replaceAll("([a-z])([A-Z])", "$1_$2")
        .replaceAll("([A-Za-z])(\\d)", "$1_$2")
        .toUpperCase(Locale.ROOT);
  }

  public static boolean isBrokenFormula(
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

  public static String locationEvidence(AnalysisLocationReport location) {
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

  public static String ruleLabel(ConditionalFormattingRuleReport rule) {
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
        fromPoi(threshold.getRangeType()), threshold.getFormula(), threshold.getValue());
  }

  private static List<String> colors(XSSFColor[] colors) {
    List<String> values = new ArrayList<>(colors.length);
    for (XSSFColor color : colors) {
      values.add(color(color));
    }
    return List.copyOf(values);
  }

  private static String color(XSSFColor color) {
    String rgb = ColorAdapter.toRgbHex(color).orElse(null);
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
          fromPoi(rule.getComparisonOperation()),
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
          fromPoi(iconFormatting.getIconSet()),
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

  /** Writes one authored differential style onto the supplied conditional-formatting rule XML. */
  public static void applyStyle(
      XSSFWorkbook workbook, CTCfRule rule, DifferentialStylePatch style) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(style, "style must not be null");

    CTDxf dxf = CTDxf.Factory.newInstance();
    applyStyle(workbook.getStylesSource(), dxf, style);
    attachStyle(workbook.getStylesSource(), rule, dxf);
  }

  /**
   * Attaches one raw differential-style XML payload to the supplied conditional-formatting rule.
   */
  public static void attachStyle(StylesTable stylesTable, CTCfRule rule, CTDxf dxf) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(dxf, "dxf must not be null");
    rule.setDxfId(putDxf(stylesTable, dxf));
  }

  /**
   * Returns the factual differential-style snapshot attached to one conditional-formatting rule.
   */
  public static @Nullable DifferentialStyleReport snapshotStyle(
      StylesTable stylesTable, CTCfRule rule) {
    return optionalSnapshotStyle(stylesTable, rule).orElse(null);
  }

  public static long putDxf(StylesTable stylesTable, CTDxf dxf) {
    return stylesTable.putDxf(dxf) - 1L;
  }

  public static Optional<CTDxf> dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0 || dxfId >= stylesTable._getDXfsSize()) {
      return Optional.empty();
    }
    return Optional.of(stylesTable.getDxfAt(Math.toIntExact(dxfId)));
  }

  public record FillSnapshot(
      @Nullable CellColorReport fillColor,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    public FillSnapshot {
      unsupportedFeatures = List.copyOf(unsupportedFeatures);
    }

    public static FillSnapshot empty() {
      return new FillSnapshot(null, List.of());
    }
  }

  public record BorderSnapshot(
      @Nullable DifferentialBorderReport border,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    public BorderSnapshot {
      unsupportedFeatures = List.copyOf(unsupportedFeatures);
    }

    public static BorderSnapshot empty() {
      return new BorderSnapshot(null, List.of());
    }
  }

  public record FontSnapshot(
      @Nullable Boolean bold,
      @Nullable Boolean italic,
      @Nullable FontHeightReport fontHeight,
      @Nullable CellColorReport fontColor,
      @Nullable Boolean underline,
      @Nullable Boolean strikeout,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    public FontSnapshot {
      unsupportedFeatures = List.copyOf(unsupportedFeatures);
    }

    public static FontSnapshot empty() {
      return new FontSnapshot(null, null, null, null, null, null, List.of());
    }

    public boolean isEmpty() {
      return java.util.stream.Stream.of(bold, italic, fontHeight, fontColor, underline, strikeout)
          .allMatch(java.util.Objects::isNull);
    }
  }

  public static void applyStyle(StylesTable stylesTable, CTDxf dxf, DifferentialStylePatch style) {
    applyNumberFormat(stylesTable, dxf, style.numberFormat().orElse(null));
    applyFont(dxf, style);
    applyFill(dxf, style.fillColor().orElse(null));
    applyBorder(dxf, style.border().orElse(null));
  }

  private static void applyNumberFormat(
      StylesTable stylesTable, CTDxf dxf, @Nullable String numberFormat) {
    if (numberFormat == null) {
      return;
    }
    int formatId = stylesTable.putNumberFormat(numberFormat);
    CTNumFmt numFmt = dxf.addNewNumFmt();
    numFmt.setNumFmtId(formatId);
    numFmt.setFormatCode(numberFormat);
  }

  private static void applyFont(CTDxf dxf, DifferentialStylePatch style) {
    if (java.util.stream.Stream.of(
            style.bold().orElse(null),
            style.italic().orElse(null),
            style.fontHeight().orElse(null),
            style.fontColor().orElse(null),
            style.underline().orElse(null),
            style.strikeout().orElse(null))
        .allMatch(Objects::isNull)) {
      return;
    }

    CTFont font = dxf.addNewFont();
    setBooleanProperty(style.bold().orElse(null), font::addNewB);
    setBooleanProperty(style.italic().orElse(null), font::addNewI);
    setBooleanProperty(style.strikeout().orElse(null), font::addNewStrike);
    if (style.fontHeight().isPresent()) {
      CTFontSize fontSize = font.addNewSz();
      fontSize.setVal(style.fontHeight().orElseThrow().points().doubleValue());
    }
    if (style.fontColor().isPresent()) {
      setColor(font.addNewColor(), style.fontColor().orElseThrow());
    }
    if (style.underline().isPresent()) {
      CTUnderlineProperty underlineProperty = font.addNewU();
      underlineProperty.setVal(
          style.underline().orElseThrow() ? STUnderlineValues.SINGLE : STUnderlineValues.NONE);
    }
  }

  private static void applyFill(CTDxf dxf, @Nullable ColorSpec fillColor) {
    if (fillColor == null) {
      return;
    }
    CTFill fill = dxf.addNewFill();
    CTPatternFill patternFill = fill.addNewPatternFill();
    patternFill.setPatternType(STPatternType.SOLID);
    setColor(patternFill.addNewFgColor(), fillColor);
  }

  private static void applyBorder(CTDxf dxf, @Nullable DifferentialBorderPatch border) {
    if (border == null) {
      return;
    }
    CTBorder ctBorder = dxf.addNewBorder();
    applyBorderSide(
        ctBorder::addNewTop, resolvedSide(border.all().orElse(null), border.top().orElse(null)));
    applyBorderSide(
        ctBorder::addNewRight,
        resolvedSide(border.all().orElse(null), border.right().orElse(null)));
    applyBorderSide(
        ctBorder::addNewBottom,
        resolvedSide(border.all().orElse(null), border.bottom().orElse(null)));
    applyBorderSide(
        ctBorder::addNewLeft, resolvedSide(border.all().orElse(null), border.left().orElse(null)));
  }

  private static void applyBorderSide(
      java.util.function.Supplier<CTBorderPr> borderPrFactory, @Nullable BorderSidePatch side) {
    Objects.requireNonNull(borderPrFactory, "borderPrFactory must not be null");
    if (side == null) {
      return;
    }
    CTBorderPr borderPr = borderPrFactory.get();
    if (side.style().isPresent()) {
      borderPr.setStyle(toCtBorderStyle(side.style().orElseThrow()));
    }
    if (side.color().isPresent()) {
      setColor(borderPr.addNewColor(), side.color().orElseThrow());
    }
  }

  private static @Nullable BorderSidePatch resolvedSide(
      @Nullable BorderSidePatch defaultSide, @Nullable BorderSidePatch explicitSide) {
    return explicitSide == null ? defaultSide : explicitSide;
  }

  private static void setBooleanProperty(
      @Nullable Boolean value, java.util.function.Supplier<CTBooleanProperty> propertyFactory) {
    if (value == null) {
      return;
    }
    CTBooleanProperty property = propertyFactory.get();
    property.setVal(value);
  }

  private static Optional<DifferentialStyleReport> optionalSnapshotStyle(
      StylesTable stylesTable, CTCfRule rule) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    if (!rule.isSetDxfId()) {
      return Optional.empty();
    }

    CTDxf dxf = dxfAt(stylesTable, rule.getDxfId()).orElse(null);
    if (dxf == null) {
      return Optional.of(
          new DifferentialStyleReport(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              List.of(ExcelConditionalFormattingUnsupportedFeature.STYLE_REFERENCE)));
    }

    FontSnapshot font = snapshotFont(dxf.isSetFont() ? dxf.getFont() : null);
    FillSnapshot fill = snapshotFill(dxf.isSetFill() ? dxf.getFill() : null);
    BorderSnapshot border = snapshotBorder(dxf.isSetBorder() ? dxf.getBorder() : null);
    return optionalStyleSnapshot(
        dxf.isSetNumFmt() ? dxf.getNumFmt().getFormatCode() : null,
        font,
        fill,
        border,
        metadataUnsupportedFeatures(dxf));
  }

  public static FillSnapshot snapshotFill(@Nullable CTFill fill) {
    if (fill == null) {
      return FillSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (hasGradientFill(fill)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    if (!fill.isSetPatternFill()) {
      return new FillSnapshot(null, normalizedUnsupportedFeatures(unsupportedFeatures));
    }

    CTPatternFill patternFill = fill.getPatternFill();
    unsupportedFeatures.addAll(patternFillUnsupportedFeatures(patternFill));
    return new FillSnapshot(
        patternForegroundColor(patternFill, unsupportedFeatures),
        normalizedUnsupportedFeatures(unsupportedFeatures));
  }

  public static boolean patternTypeIsUnsupported(CTPatternFill patternFill) {
    return patternFill.isSetPatternType()
        && patternFill.getPatternType() != STPatternType.SOLID
        && patternFill.getPatternType() != STPatternType.NONE;
  }

  public static @Nullable CellColorReport patternForegroundColor(
      CTPatternFill patternFill,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    CellColorReport fillColor = optionalPatternForegroundColor(patternFill).orElse(null);
    if (fillColor == null
        && (patternTypeIsUnsupported(patternFill) || patternFill.isSetFgColor())) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    return fillColor;
  }

  public static boolean hasComplexBorderFeatures(CTBorder border) {
    return java.util.stream.Stream.of(
            border.isSetDiagonal(),
            border.isSetVertical(),
            border.isSetHorizontal(),
            border.isSetStart(),
            border.isSetEnd(),
            border.isSetDiagonalDown(),
            border.isSetDiagonalUp())
        .anyMatch(Boolean.TRUE::equals);
  }

  public static boolean hasUnsupportedSideReference(
      CTBorder border,
      @Nullable DifferentialBorderSideReport top,
      @Nullable DifferentialBorderSideReport right,
      @Nullable DifferentialBorderSideReport bottom,
      @Nullable DifferentialBorderSideReport left) {
    return java.util.stream.Stream.of(
            border.isSetTop() && top == null,
            border.isSetRight() && right == null,
            border.isSetBottom() && bottom == null,
            border.isSetLeft() && left == null)
        .anyMatch(Boolean.TRUE::equals);
  }

  public static @Nullable DifferentialBorderReport borderValue(
      @Nullable DifferentialBorderSideReport top,
      @Nullable DifferentialBorderSideReport right,
      @Nullable DifferentialBorderSideReport bottom,
      @Nullable DifferentialBorderSideReport left) {
    return optionalBorderValue(top, right, bottom, left).orElse(null);
  }

  public static @Nullable DifferentialBorderSideReport snapshotBorderSide(
      @Nullable CTBorderPr side) {
    return optionalSnapshotBorderSide(side).orElse(null);
  }

  public static @Nullable Boolean booleanProperty(@Nullable CTBooleanProperty property) {
    return property == null ? null : !property.isSetVal() || property.getVal();
  }

  public static boolean underline(@Nullable CTUnderlineProperty property) {
    if (property == null || !property.isSetVal()) {
      return true;
    }
    return property.getVal() != STUnderlineValues.NONE;
  }

  public static boolean hasUnsupportedFontAttributes(CTFont font) {
    return java.util.stream.IntStream.of(
            font.sizeOfNameArray(),
            font.sizeOfCharsetArray(),
            font.sizeOfFamilyArray(),
            font.sizeOfOutlineArray(),
            font.sizeOfShadowArray(),
            font.sizeOfCondenseArray(),
            font.sizeOfExtendArray(),
            font.sizeOfVertAlignArray(),
            font.sizeOfSchemeArray())
        .anyMatch(size -> size > 0);
  }

  public static List<ExcelConditionalFormattingUnsupportedFeature> metadataUnsupportedFeatures(
      CTDxf dxf) {
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (dxf.isSetAlignment()) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT);
    }
    if (dxf.isSetProtection()) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.PROTECTION);
    }
    return List.copyOf(unsupportedFeatures);
  }

  private static Optional<DifferentialStyleReport> optionalStyleSnapshot(
      @Nullable String numberFormat,
      FontSnapshot font,
      FillSnapshot fill,
      BorderSnapshot border,
      List<ExcelConditionalFormattingUnsupportedFeature> metadataUnsupportedFeatures) {
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        normalizedUnsupportedFeatures(
            font.unsupportedFeatures(),
            fill.unsupportedFeatures(),
            border.unsupportedFeatures(),
            metadataUnsupportedFeatures);
    if (java.util.stream.Stream.of(numberFormat, fill.fillColor(), border.border())
            .allMatch(Objects::isNull)
        && font.isEmpty()
        && unsupportedFeatures.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(
        new DifferentialStyleReport(
            Optional.ofNullable(numberFormat),
            Optional.ofNullable(font.bold()),
            Optional.ofNullable(font.italic()),
            Optional.ofNullable(font.fontHeight()),
            Optional.ofNullable(font.fontColor()),
            Optional.ofNullable(font.underline()),
            Optional.ofNullable(font.strikeout()),
            Optional.ofNullable(fill.fillColor()),
            Optional.ofNullable(border.border()),
            unsupportedFeatures));
  }

  private static FontSnapshot snapshotFont(@Nullable CTFont font) {
    if (font == null) {
      return FontSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    CellColorReport fontColor = fontColor(font, unsupportedFeatures);
    if (hasUnsupportedFontAttributes(font)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES);
    }

    return new FontSnapshot(
        booleanProperty(font.sizeOfBArray() > 0 ? font.getBArray(0) : null),
        booleanProperty(font.sizeOfIArray() > 0 ? font.getIArray(0) : null),
        fontHeight(font).orElse(null),
        fontColor,
        font.sizeOfUArray() > 0 ? underline(font.getUArray(0)) : null,
        booleanProperty(font.sizeOfStrikeArray() > 0 ? font.getStrikeArray(0) : null),
        normalizedUnsupportedFeatures(unsupportedFeatures));
  }

  private static Optional<FontHeightReport> fontHeight(CTFont font) {
    if (font.sizeOfSzArray() == 0) {
      return Optional.empty();
    }
    return Optional.of(
        FontHeightReport.fromPoints(BigDecimal.valueOf(font.getSzArray(0).getVal())));
  }

  private static @Nullable CellColorReport fontColor(
      CTFont font, List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    CellColorReport fontColor = optionalFontColor(font).orElse(null);
    if (fontColor == null && font.sizeOfColorArray() > 0) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES);
    }
    return fontColor;
  }

  private static boolean hasGradientFill(CTFill fill) {
    return fill.xmlText().contains("gradientFill");
  }

  private static BorderSnapshot snapshotBorder(@Nullable CTBorder border) {
    if (border == null) {
      return BorderSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (hasComplexBorderFeatures(border)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY);
    }

    DifferentialBorderSideReport top = borderSide(border.isSetTop(), border::getTop);
    DifferentialBorderSideReport right = borderSide(border.isSetRight(), border::getRight);
    DifferentialBorderSideReport bottom = borderSide(border.isSetBottom(), border::getBottom);
    DifferentialBorderSideReport left = borderSide(border.isSetLeft(), border::getLeft);

    if (hasUnsupportedSideReference(border, top, right, bottom, left)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY);
    }

    return new BorderSnapshot(
        borderValue(top, right, bottom, left), normalizedUnsupportedFeatures(unsupportedFeatures));
  }

  private static List<ExcelConditionalFormattingUnsupportedFeature> patternFillUnsupportedFeatures(
      CTPatternFill patternFill) {
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (patternFill.isSetBgColor()) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_BACKGROUND_COLOR);
    }
    if (patternTypeIsUnsupported(patternFill)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    return List.copyOf(unsupportedFeatures);
  }

  private static @Nullable DifferentialBorderSideReport borderSide(
      boolean present, java.util.function.Supplier<CTBorderPr> sideSupplier) {
    return present ? snapshotBorderSide(sideSupplier.get()) : null;
  }

  private static Optional<CellColorReport> optionalFontColor(CTFont font) {
    return font.sizeOfColorArray() == 0 ? Optional.empty() : optionalColor(font.getColorArray(0));
  }

  private static Optional<CellColorReport> optionalPatternForegroundColor(
      CTPatternFill patternFill) {
    if (patternTypeIsUnsupported(patternFill) || !patternFill.isSetFgColor()) {
      return Optional.empty();
    }
    return optionalColor(patternFill.getFgColor());
  }

  private static Optional<DifferentialBorderReport> optionalBorderValue(
      @Nullable DifferentialBorderSideReport top,
      @Nullable DifferentialBorderSideReport right,
      @Nullable DifferentialBorderSideReport bottom,
      @Nullable DifferentialBorderSideReport left) {
    if (top == null && right == null && bottom == null && left == null) {
      return Optional.empty();
    }
    return Optional.of(
        new DifferentialBorderReport(
            Optional.empty(),
            Optional.ofNullable(top),
            Optional.ofNullable(right),
            Optional.ofNullable(bottom),
            Optional.ofNullable(left)));
  }

  private static Optional<DifferentialBorderSideReport> optionalSnapshotBorderSide(
      @Nullable CTBorderPr side) {
    if (side == null) {
      return Optional.empty();
    }
    if (!side.isSetStyle() && !side.isSetColor()) {
      return Optional.empty();
    }
    Optional<ExcelBorderStyle> style =
        side.isSetStyle()
            ? Optional.of(fromCtBorderStyle(side.getStyle().intValue()))
            : Optional.empty();
    Optional<CellColorReport> color =
        side.isSetColor() ? optionalColor(side.getColor()) : Optional.empty();
    if (side.isSetColor() && color.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(new DifferentialBorderSideReport(style, color));
  }

  @SafeVarargs
  private static List<ExcelConditionalFormattingUnsupportedFeature> normalizedUnsupportedFeatures(
      List<ExcelConditionalFormattingUnsupportedFeature>... groups) {
    java.util.Set<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new LinkedHashSet<>();
    for (List<ExcelConditionalFormattingUnsupportedFeature> group : groups) {
      unsupportedFeatures.addAll(group);
    }
    return List.copyOf(unsupportedFeatures);
  }

  /**
   * Writes one owned workbook color without collapsing theme, indexed, or tint semantics to RGB.
   */
  public static void setColor(CTColor target, ColorSpec color) {
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(color, "color must not be null");
    switch (color) {
      case ColorSpec.Rgb rgb -> {
        target.setRgb(argbBytes(rgb));
        rgb.tint().ifPresent(target::setTint);
      }
      case ColorSpec.Theme theme -> {
        target.setTheme(theme.theme());
        theme.tint().ifPresent(target::setTint);
      }
      case ColorSpec.Indexed indexed -> {
        target.setIndexed(indexed.indexed());
        indexed.tint().ifPresent(target::setTint);
      }
    }
  }

  /** Reads one OOXML color into the owned color vocabulary, preserving tint and reference kind. */
  public static Optional<CellColorReport> optionalColor(@Nullable CTColor color) {
    if (color == null) {
      return Optional.empty();
    }
    Optional<Double> tint = color.isSetTint() ? Optional.of(color.getTint()) : Optional.empty();
    if (color.isSetRgb()) {
      return rgbHex(color.getRgb()).map(value -> new CellColorReport.Rgb(new RgbHex(value), tint));
    }
    if (color.isSetTheme()) {
      return Optional.of(new CellColorReport.Theme(Math.toIntExact(color.getTheme()), tint));
    }
    if (color.isSetIndexed()) {
      return Optional.of(new CellColorReport.Indexed(Math.toIntExact(color.getIndexed()), tint));
    }
    return Optional.empty();
  }

  private static byte[] argbBytes(ColorSpec.Rgb color) {
    String rgb = color.rgb().value();
    return new byte[] {
      (byte) 0xFF,
      (byte) Integer.parseInt(rgb.substring(1, 3), 16),
      (byte) Integer.parseInt(rgb.substring(3, 5), 16),
      (byte) Integer.parseInt(rgb.substring(5, 7), 16)
    };
  }

  private static Optional<String> rgbHex(byte[] rgb) {
    if (rgb.length == 4) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[1] & 0xFF, rgb[2] & 0xFF, rgb[3] & 0xFF));
    }
    if (rgb.length == 3) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
    }
    return Optional.empty();
  }

  private static final Map<ExcelBorderStyle, STBorderStyle.Enum> TO_CT_BORDER_STYLE =
      EnumMappings.exactEnumMap(
          ExcelBorderStyle.class,
          "Conditional-formatting CT border-style mapping",
          Map.ofEntries(
              Map.entry(ExcelBorderStyle.NONE, STBorderStyle.NONE),
              Map.entry(ExcelBorderStyle.THIN, STBorderStyle.THIN),
              Map.entry(ExcelBorderStyle.MEDIUM, STBorderStyle.MEDIUM),
              Map.entry(ExcelBorderStyle.DASHED, STBorderStyle.DASHED),
              Map.entry(ExcelBorderStyle.DOTTED, STBorderStyle.DOTTED),
              Map.entry(ExcelBorderStyle.THICK, STBorderStyle.THICK),
              Map.entry(ExcelBorderStyle.DOUBLE, STBorderStyle.DOUBLE),
              Map.entry(ExcelBorderStyle.HAIR, STBorderStyle.HAIR),
              Map.entry(ExcelBorderStyle.MEDIUM_DASHED, STBorderStyle.MEDIUM_DASHED),
              Map.entry(ExcelBorderStyle.DASH_DOT, STBorderStyle.DASH_DOT),
              Map.entry(ExcelBorderStyle.MEDIUM_DASH_DOT, STBorderStyle.MEDIUM_DASH_DOT),
              Map.entry(ExcelBorderStyle.DASH_DOT_DOT, STBorderStyle.DASH_DOT_DOT),
              Map.entry(ExcelBorderStyle.MEDIUM_DASH_DOT_DOT, STBorderStyle.MEDIUM_DASH_DOT_DOT),
              Map.entry(ExcelBorderStyle.SLANTED_DASH_DOT, STBorderStyle.SLANT_DASH_DOT)));

  private static final Map<Integer, ExcelBorderStyle> FROM_CT_BORDER_STYLE =
      Map.ofEntries(
          Map.entry(STBorderStyle.INT_NONE, ExcelBorderStyle.NONE),
          Map.entry(STBorderStyle.INT_THIN, ExcelBorderStyle.THIN),
          Map.entry(STBorderStyle.INT_MEDIUM, ExcelBorderStyle.MEDIUM),
          Map.entry(STBorderStyle.INT_DASHED, ExcelBorderStyle.DASHED),
          Map.entry(STBorderStyle.INT_DOTTED, ExcelBorderStyle.DOTTED),
          Map.entry(STBorderStyle.INT_THICK, ExcelBorderStyle.THICK),
          Map.entry(STBorderStyle.INT_DOUBLE, ExcelBorderStyle.DOUBLE),
          Map.entry(STBorderStyle.INT_HAIR, ExcelBorderStyle.HAIR),
          Map.entry(STBorderStyle.INT_MEDIUM_DASHED, ExcelBorderStyle.MEDIUM_DASHED),
          Map.entry(STBorderStyle.INT_DASH_DOT, ExcelBorderStyle.DASH_DOT),
          Map.entry(STBorderStyle.INT_MEDIUM_DASH_DOT, ExcelBorderStyle.MEDIUM_DASH_DOT),
          Map.entry(STBorderStyle.INT_DASH_DOT_DOT, ExcelBorderStyle.DASH_DOT_DOT),
          Map.entry(STBorderStyle.INT_MEDIUM_DASH_DOT_DOT, ExcelBorderStyle.MEDIUM_DASH_DOT_DOT),
          Map.entry(STBorderStyle.INT_SLANT_DASH_DOT, ExcelBorderStyle.SLANTED_DASH_DOT));

  public static STBorderStyle.Enum toCtBorderStyle(ExcelBorderStyle style) {
    return EnumMappings.requireMappedValue(TO_CT_BORDER_STYLE, style, "GridGrind border style");
  }

  public static ExcelBorderStyle fromCtBorderStyle(int styleCode) {
    ExcelBorderStyle resolved = FROM_CT_BORDER_STYLE.get(styleCode);
    if (resolved == null) {
      throw new IllegalArgumentException("Unsupported CT border style: " + styleCode);
    }
    return resolved;
  }

  /** Returns the POI conditional-formatting operator constant for one protocol operator. */
  public static byte toPoi(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> ComparisonOperator.BETWEEN;
      case NOT_BETWEEN -> ComparisonOperator.NOT_BETWEEN;
      case EQUAL -> ComparisonOperator.EQUAL;
      case NOT_EQUAL -> ComparisonOperator.NOT_EQUAL;
      case GREATER_THAN -> ComparisonOperator.GT;
      case LESS_THAN -> ComparisonOperator.LT;
      case GREATER_OR_EQUAL -> ComparisonOperator.GE;
      case LESS_OR_EQUAL -> ComparisonOperator.LE;
    };
  }

  /** Returns the protocol operator represented by one POI conditional-formatting constant. */
  public static ExcelComparisonOperator fromPoi(int comparisonOperator) {
    return switch (comparisonOperator) {
      case ComparisonOperator.BETWEEN -> ExcelComparisonOperator.BETWEEN;
      case ComparisonOperator.NOT_BETWEEN -> ExcelComparisonOperator.NOT_BETWEEN;
      case ComparisonOperator.EQUAL -> ExcelComparisonOperator.EQUAL;
      case ComparisonOperator.NOT_EQUAL -> ExcelComparisonOperator.NOT_EQUAL;
      case ComparisonOperator.GT -> ExcelComparisonOperator.GREATER_THAN;
      case ComparisonOperator.LT -> ExcelComparisonOperator.LESS_THAN;
      case ComparisonOperator.GE -> ExcelComparisonOperator.GREATER_OR_EQUAL;
      case ComparisonOperator.LE -> ExcelComparisonOperator.LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported Apache POI conditional-formatting comparison operator: "
                  + comparisonOperator);
    };
  }

  public static ExcelConditionalFormattingIconSet fromPoi(
      IconMultiStateFormatting.IconSet iconSet) {
    Objects.requireNonNull(iconSet, "iconSet must not be null");
    return ExcelConditionalFormattingIconSet.valueOf(iconSet.name());
  }

  public static IconMultiStateFormatting.IconSet toPoi(ExcelConditionalFormattingIconSet iconSet) {
    Objects.requireNonNull(iconSet, "iconSet must not be null");
    return IconMultiStateFormatting.IconSet.valueOf(iconSet.name());
  }

  public static ExcelConditionalFormattingThresholdType fromPoi(
      ConditionalFormattingThreshold.RangeType type) {
    Objects.requireNonNull(type, "type must not be null");
    return ExcelConditionalFormattingThresholdType.valueOf(type.name());
  }

  public static ConditionalFormattingThreshold.RangeType toPoi(
      ExcelConditionalFormattingThresholdType type) {
    Objects.requireNonNull(type, "type must not be null");
    return ConditionalFormattingThreshold.RangeType.valueOf(type.name());
  }
}
