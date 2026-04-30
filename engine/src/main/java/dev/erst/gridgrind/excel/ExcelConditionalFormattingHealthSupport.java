package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Derives health findings from canonical conditional-formatting snapshots. */
final class ExcelConditionalFormattingHealthSupport {
  List<WorkbookAnalysis.AnalysisFinding> conditionalFormattingHealthFindings(
      String sheetName, XSSFSheet sheet, List<ExcelConditionalFormattingBlockSnapshot> blocks) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    List<PriorityContext> priorities = new ArrayList<>();
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(sheet.getWorkbook());

    for (ExcelConditionalFormattingBlockSnapshot block : blocks) {
      if (block.ranges().isEmpty() || hasInvalidRanges(block.ranges())) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                AnalysisFindingCode.CONDITIONAL_FORMATTING_EMPTY_RANGE,
                AnalysisSeverity.WARNING,
                "Conditional-formatting block targets an empty or invalid range",
                "Conditional-formatting block has no valid target ranges.",
                blockLocation(sheetName, block.ranges()),
                List.copyOf(block.ranges())));
      }

      for (ExcelConditionalFormattingRuleSnapshot rule : block.rules()) {
        priorities.add(priorityContext(sheetName, block, rule));
        switch (rule) {
          case ExcelConditionalFormattingRuleSnapshot.FormulaRule formulaRule -> {
            if (ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
                evaluationWorkbook, sheetIndex, formulaRule.formula())) {
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
                      AnalysisSeverity.ERROR,
                      "Conditional-formatting formula is invalid",
                      "Formula rule could not be parsed for conditional-formatting evaluation.",
                      blockLocation(sheetName, block.ranges()),
                      List.of(formulaRule.formula())));
            }
          }
          case ExcelConditionalFormattingRuleSnapshot.CellValueRule cellValueRule -> {
            if (ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
                    evaluationWorkbook, sheetIndex, cellValueRule.formula1())
                || (cellValueRule.formula2() != null
                    && ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
                        evaluationWorkbook, sheetIndex, cellValueRule.formula2()))) {
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
                      AnalysisSeverity.ERROR,
                      "Conditional-formatting operand formula is invalid",
                      "Cell-value rule operands could not be parsed for conditional-formatting evaluation.",
                      blockLocation(sheetName, block.ranges()),
                      cellValueRuleEvidence(cellValueRule)));
            }
          }
          case ExcelConditionalFormattingRuleSnapshot.UnsupportedRule unsupportedRule ->
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      AnalysisFindingCode.CONDITIONAL_FORMATTING_UNSUPPORTED_RULE,
                      AnalysisSeverity.WARNING,
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

  private static boolean hasInvalidRanges(List<String> ranges) {
    for (String range : ranges) {
      if (ExcelSheetStructureSupport.parseRangeOrNull(range) == null) {
        return true;
      }
    }
    return false;
  }

  private static PriorityContext priorityContext(
      String sheetName,
      ExcelConditionalFormattingBlockSnapshot block,
      ExcelConditionalFormattingRuleSnapshot rule) {
    return new PriorityContext(
        rule.priority(),
        new BlockRuleContext(
            blockLocation(sheetName, block.ranges()),
            block.ranges(),
            ExcelConditionalFormattingSnapshotSupport.ruleLabel(rule)));
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
      String sheetName, List<String> ranges) {
    if (!ranges.isEmpty()) {
      return new WorkbookAnalysis.AnalysisLocation.Range(sheetName, ranges.getFirst());
    }
    return new WorkbookAnalysis.AnalysisLocation.Sheet(sheetName);
  }

  private record BlockRuleContext(
      WorkbookAnalysis.AnalysisLocation location, List<String> ranges, String label) {}

  private record PriorityContext(int priority, BlockRuleContext context) {}
}
