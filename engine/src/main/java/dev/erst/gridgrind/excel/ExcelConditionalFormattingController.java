package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
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
  void setConditionalFormatting(
      XSSFSheet sheet,
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges,
      dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec formatting) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(formatting, "formatting must not be null");
    authoring.setConditionalFormatting(sheet, ranges, formatting);
  }

  /** Returns the number of conditional-formatting blocks currently present on one sheet. */
  int conditionalFormattingBlockCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.getSheetConditionalFormatting().getNumConditionalFormattings();
  }

  /** Returns derived conditional-formatting health findings for one sheet. */
  List<AnalysisFindingReport> conditionalFormattingHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    return health.conditionalFormattingHealthFindings(
        sheetName, sheet, conditionalFormatting(sheet));
  }

  /** Converts one loaded POI conditional-formatting rule into the matching GridGrind snapshot. */
  static ConditionalFormattingRuleReport toSnapshot(
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
  static String locationEvidence(AnalysisLocationReport location) {
    return ExcelConditionalFormattingSnapshotSupport.locationEvidence(location);
  }

  /** Returns the stable label used in priority-collision evidence for one snapshot rule family. */
  static String ruleLabel(ConditionalFormattingRuleReport rule) {
    return ExcelConditionalFormattingSnapshotSupport.ruleLabel(rule);
  }

  /** Applies the explicit bound-sheet range operation. */
  void clearConditionalFormatting(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    authoring.clearConditionalFormatting(sheet);
  }

  /** Applies the explicit bound-sheet range operation. */
  void clearConditionalFormatting(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    authoring.clearConditionalFormatting(sheet, ranges);
  }

  /** Applies the explicit bound-sheet range operation. */
  List<ConditionalFormattingEntryReport> conditionalFormatting(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return snapshots.conditionalFormatting(sheet);
  }

  /** Applies the explicit bound-sheet range operation. */
  List<ConditionalFormattingEntryReport> conditionalFormatting(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return snapshots.conditionalFormatting(sheet, ranges);
  }
}
