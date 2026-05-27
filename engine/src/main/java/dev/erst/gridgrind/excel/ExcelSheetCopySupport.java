package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

/** Rebuilds and retargets the sheet-local structures that GridGrind can reproduce safely. */
final class ExcelSheetCopySupport {
  private ExcelSheetCopySupport() {}

  static void repairComments(
      List<WorkbookSheetResult.CellComment> expectedComments, ExcelSheet targetSheet) {
    if (expectedComments.isEmpty()) {
      return;
    }
    // cloneSheet can leave copied comments looking correct in-memory while reopening later
    // without a stable client anchor. Rewriting the copied comments makes the persisted
    // VML-backed comment state authoritative again.
    ExcelSheetCommentRepairSupport commentRepairSupport =
        new ExcelSheetCommentRepairSupport(targetSheet.xssfSheet());
    commentRepairSupport.replaceComments(
        ExcelSheetCommentRepairSupport.commentRewriteSnapshots(expectedComments));
  }

  static void repairPrintLayout(ExcelPrintLayout expectedPrintLayout, ExcelSheet targetSheet) {
    if (expectedPrintLayout.equals(targetSheet.layout().printLayout())) {
      return;
    }
    targetSheet.layout().setPrintLayout(expectedPrintLayout);
  }

  static void replaceLocalNamedRanges(
      ExcelWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<ExcelNamedRangeSnapshot> localNamedRanges) {
    ExcelSheetCopyNamedRangeSupport.replaceLocalNamedRanges(
        workbook, sourceSheetName, targetSheetName, localNamedRanges);
  }

  static void deleteLocalNamedRanges(ExcelWorkbook workbook, String sheetName) {
    ExcelSheetCopyNamedRangeSupport.deleteLocalNamedRanges(workbook, sheetName);
  }

  static void replaceDataValidations(
      List<CTDataValidation> validations,
      XSSFSheet targetPoiSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    ExcelSheetCopyValidationSupport.replaceDataValidations(
        validations, targetPoiSheet, workbook, sourceSheetName, newSheetName);
  }

  static void replaceConditionalFormatting(
      List<ExcelConditionalFormattingBlockDefinition> blocks,
      ExcelSheet targetSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    ExcelSheetCopyConditionalFormattingSupport.replaceConditionalFormatting(
        blocks, targetSheet, workbook, sourceSheetName, newSheetName);
  }

  static void replaceTables(
      ExcelWorkbook workbook, String targetSheetName, List<ExcelTableSnapshot> tables) {
    ExcelSheetCopyTableSupport.replaceTables(workbook, targetSheetName, tables);
  }

  static void replaceAutofilter(
      Optional<ExcelAutofilterSnapshot.SheetOwned> sheetAutofilter, ExcelSheet targetSheet) {
    ExcelSheetCopyAutofilterSupport.replaceAutofilter(sheetAutofilter, targetSheet);
  }

  static void copyProtection(XSSFSheet sourceSheet, XSSFSheet targetSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    if (!sourceSheet.getProtect()) {
      return;
    }
    CTSheetProtection copiedProtection =
        (CTSheetProtection) sourceSheet.getCTWorksheet().getSheetProtection().copy();
    targetSheet.getCTWorksheet().setSheetProtection(copiedProtection);
  }

  static List<ExcelNamedRangeSnapshot> copyableLocalNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    return ExcelSheetCopyNamedRangeSupport.copyableLocalNames(namedRanges, sheetName);
  }

  static List<ExcelNamedRangeSnapshot.RangeSnapshot> copyableLocalRangeNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    return ExcelSheetCopyNamedRangeSupport.copyableLocalRangeNames(namedRanges, sheetName);
  }

  static void requireNoUncopyableLocalNamedRanges(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    ExcelSheetCopyNamedRangeSupport.requireNoUncopyableLocalNamedRanges(namedRanges, sheetName);
  }

  static void requireNoTables(XSSFSheet sheet, String sheetName) {
    ExcelSheetCopyTableSupport.requireNoTables(sheet, sheetName);
  }

  static List<ExcelConditionalFormattingBlockDefinition> supportedConditionalFormatting(
      List<ExcelConditionalFormattingBlockSnapshot> blocks, String sourceSheetName) {
    return ExcelSheetCopyConditionalFormattingSupport.supportedConditionalFormatting(
        blocks, sourceSheetName);
  }

  static List<CTDataValidation> copiedDataValidations(XSSFSheet sourcePoiSheet) {
    return ExcelSheetCopyValidationSupport.copiedDataValidations(sourcePoiSheet);
  }

  static ExcelConditionalFormattingRule copyableRule(
      ExcelConditionalFormattingRuleSnapshot rule, String sourceSheetName) {
    return ExcelSheetCopyConditionalFormattingSupport.copyableRule(rule, sourceSheetName);
  }

  static Optional<ExcelDifferentialStyle> copyableStyle(
      @Nullable ExcelDifferentialStyleSnapshot style, String sourceSheetName) {
    return ExcelSheetCopyConditionalFormattingSupport.copyableStyle(style, sourceSheetName);
  }

  static Optional<ExcelAutofilterSnapshot.SheetOwned> sheetOwnedAutofilter(
      List<ExcelAutofilterSnapshot> autofilters) {
    return ExcelSheetCopyAutofilterSupport.sheetOwnedAutofilter(autofilters);
  }

  static Optional<String> sheetOwnedAutofilterRange(List<ExcelAutofilterSnapshot> autofilters) {
    return ExcelSheetCopyAutofilterSupport.sheetOwnedAutofilterRange(autofilters);
  }

  static ExcelAutofilterSortCondition copyableSortCondition(
      ExcelAutofilterSortConditionSnapshot condition) {
    return ExcelSheetCopyAutofilterSupport.copyableSortCondition(condition);
  }

  static List<ExcelTableSnapshot> tablesOnSheet(XSSFSheet sourcePoiSheet) {
    return ExcelSheetCopyTableSupport.tablesOnSheet(sourcePoiSheet);
  }

  static String retargetFormula(
      ExcelWorkbook workbook,
      String formula,
      FormulaType formulaType,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    return ExcelFormulaSheetRenameSupport.renameSheet(
        workbook.xssfWorkbook(),
        formula,
        formulaType,
        targetSheetIndex,
        sourceSheetName,
        newSheetName);
  }

  static Optional<String> retargetOptionalFormula(
      ExcelWorkbook workbook,
      Optional<String> formula,
      FormulaType formulaType,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    Objects.requireNonNull(formula, "formula must not be null");
    return formula.map(
        value ->
            retargetFormula(
                workbook, value, formulaType, targetSheetIndex, sourceSheetName, newSheetName));
  }
}
