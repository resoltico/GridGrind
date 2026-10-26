package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Derives health findings from canonical conditional-formatting snapshots. */
final class ExcelConditionalFormattingHealthSupport {
  List<WorkbookAnalysis.AnalysisFinding> conditionalFormattingHealthFindings(
      String sheetName, XSSFSheet sheet, List<ConditionalFormattingEntryReport> blocks) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
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
      List<WorkbookAnalysis.AnalysisFinding> findings,
      String sheetName,
      ConditionalFormattingEntryReport block) {
    if (block.ranges().isEmpty() || hasInvalidRanges(block.ranges())) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
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
      List<WorkbookAnalysis.AnalysisFinding> findings,
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
              new WorkbookAnalysis.AnalysisFinding(
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
      List<WorkbookAnalysis.AnalysisFinding> findings,
      XSSFEvaluationWorkbook evaluationWorkbook,
      int sheetIndex,
      String sheetName,
      ConditionalFormattingEntryReport block,
      ConditionalFormattingRuleReport.FormulaRule formulaRule) {
    if (!ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
        evaluationWorkbook, sheetIndex, formulaRule.formula())) {
      return;
    }
    findings.add(
        new WorkbookAnalysis.AnalysisFinding(
            AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
            AnalysisSeverity.ERROR,
            "Conditional-formatting formula is invalid",
            "Formula rule could not be parsed for conditional-formatting evaluation.",
            blockLocation(sheetName, block.ranges()),
            List.of(formulaRule.formula())));
  }

  private static void addCellValueRuleFinding(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      XSSFEvaluationWorkbook evaluationWorkbook,
      int sheetIndex,
      String sheetName,
      ConditionalFormattingEntryReport block,
      ConditionalFormattingRuleReport.CellValueRule cellValueRule) {
    if (!hasBrokenCellValueOperand(evaluationWorkbook, sheetIndex, cellValueRule)) {
      return;
    }
    findings.add(
        new WorkbookAnalysis.AnalysisFinding(
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
    return ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
            evaluationWorkbook, sheetIndex, cellValueRule.formula1())
        || (cellValueRule.formula2().isPresent()
            && ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
                evaluationWorkbook, sheetIndex, cellValueRule.formula2().orElseThrow()));
  }

  private static boolean hasInvalidRanges(
      List<dev.erst.gridgrind.core.vocabulary.RangeLabel> ranges) {
    for (var range : ranges) {
      if (ExcelSheetStructureSupport.parseOptionalRange(range.value()).isEmpty()) {
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
            blockLocation(sheetName, block.ranges()),
            block.ranges(),
            ExcelConditionalFormattingSnapshotSupport.ruleLabel(rule)));
  }

  private static List<String> cellValueRuleEvidence(
      ConditionalFormattingRuleReport.CellValueRule cellValueRule) {
    if (cellValueRule.formula2().isEmpty()) {
      return List.of(cellValueRule.formula1());
    }
    return List.of(cellValueRule.formula1(), cellValueRule.formula2().orElseThrow());
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
            .map(
                context ->
                    context.label()
                        + "@"
                        + ExcelConditionalFormattingSnapshotSupport.locationEvidence(
                            context.location()))
            .toList();
    return new WorkbookAnalysis.AnalysisFinding(
        AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
        AnalysisSeverity.WARNING,
        "Conditional-formatting priorities collide",
        "Conditional-formatting priorities must be positive and unique within a sheet.",
        contexts.getFirst().context().location(),
        evidence);
  }

  private static WorkbookAnalysis.AnalysisLocation blockLocation(
      String sheetName, List<dev.erst.gridgrind.core.vocabulary.RangeLabel> ranges) {
    if (!ranges.isEmpty()) {
      return new WorkbookAnalysis.AnalysisLocation.Range(sheetName, ranges.getFirst().value());
    }
    return new WorkbookAnalysis.AnalysisLocation.Sheet(sheetName);
  }

  private record BlockRuleContext(
      WorkbookAnalysis.AnalysisLocation location,
      List<dev.erst.gridgrind.core.vocabulary.RangeLabel> ranges,
      String label) {
    private BlockRuleContext {
      ranges = List.copyOf(ranges);
    }
  }

  private record PriorityContext(int priority, BlockRuleContext context) {}
}
