package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;

/** Reads, writes, and analyzes conditional-formatting blocks on one XSSF sheet. */
final class ExcelConditionalFormattingController {
  private final ExcelConditionalFormattingAuthoringSupport authoring =
      new ExcelConditionalFormattingAuthoringSupport();
  private final ExcelConditionalFormattingSnapshotSupport snapshots =
      new ExcelConditionalFormattingSnapshotSupport();
  private final ExcelConditionalFormattingHealthSupport health =
      new ExcelConditionalFormattingHealthSupport();

  /** Creates or replaces one logical conditional-formatting block on the target sheet. */
  void setConditionalFormatting(XSSFSheet sheet, ExcelConditionalFormattingBlockDefinition block) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(block, "block must not be null");
    authoring.setConditionalFormatting(sheet, block);
  }

  /** Removes conditional-formatting blocks on one sheet according to the provided selection. */
  void clearConditionalFormatting(XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    authoring.clearConditionalFormatting(sheet, selection);
  }

  /** Returns factual conditional-formatting blocks for one sheet and range selection. */
  List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return snapshots.conditionalFormatting(sheet, selection);
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
    return health.conditionalFormattingHealthFindings(
        sheetName, sheet, conditionalFormatting(sheet, new ExcelRangeSelection.All()));
  }

  /** Converts one loaded POI conditional-formatting rule into the matching GridGrind snapshot. */
  static ExcelConditionalFormattingRuleSnapshot toSnapshot(
      XSSFSheet sheet, XSSFConditionalFormattingRule rule, CTCfRule ctRule) {
    return ExcelConditionalFormattingSnapshotSupport.toSnapshot(sheet, rule, ctRule);
  }

  /** Returns the stable unsupported-rule kind label for one unmodeled conditional-format family. */
  static String unsupportedKind(ConditionType conditionType, CTCfRule ctRule) {
    return ExcelConditionalFormattingSnapshotSupport.unsupportedKind(conditionType, ctRule);
  }

  /** Normalizes one raw conditional-format family name into GridGrind's stable uppercase kind. */
  static String normalizedUnsupportedKind(String rawKind) {
    return ExcelConditionalFormattingSnapshotSupport.normalizedUnsupportedKind(rawKind);
  }

  /** Reports whether one conditional-formatting formula is syntactically broken for evaluation. */
  static boolean isBrokenFormula(
      XSSFEvaluationWorkbook evaluationWorkbook, int sheetIndex, String formula) {
    return ExcelConditionalFormattingSnapshotSupport.isBrokenFormula(
        evaluationWorkbook, sheetIndex, formula);
  }

  /** Formats one analysis location into the evidence string used by collision findings. */
  static String locationEvidence(WorkbookAnalysis.AnalysisLocation location) {
    return ExcelConditionalFormattingSnapshotSupport.locationEvidence(location);
  }

  /** Returns the stable label used in priority-collision evidence for one snapshot rule family. */
  static String ruleLabel(ExcelConditionalFormattingRuleSnapshot rule) {
    return ExcelConditionalFormattingSnapshotSupport.ruleLabel(rule);
  }
}
