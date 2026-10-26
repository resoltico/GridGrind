package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleReport;
import java.util.List;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;

/** Preserves existing rule XML and retargets formulas when copying a sheet. */
final class ExcelSheetCopyConditionalFormattingSupport {
  private ExcelSheetCopyConditionalFormattingSupport() {}

  static List<CTConditionalFormatting> snapshot(XSSFSheet sheet, String sheetName) {
    var blocks =
        new ExcelConditionalFormattingController()
            .conditionalFormatting(sheet, new ExcelRangeSelection.All());
    for (var block : blocks) {
      for (var rule : block.rules()) {
        switch (rule) {
          case ConditionalFormattingRuleReport.UnsupportedRule unsupported ->
              throw new IllegalArgumentException(
                  "cannot copy sheet '"
                      + sheetName
                      + "': unsupported conditional-formatting rule '"
                      + unsupported.kind()
                      + "' is not copyable");
          case ConditionalFormattingRuleReport.FormulaRule value ->
              requireSupportedStyle(value.style().orElse(null), sheetName);
          case ConditionalFormattingRuleReport.CellValueRule value ->
              requireSupportedStyle(value.style().orElse(null), sheetName);
          case ConditionalFormattingRuleReport.Top10Rule value ->
              requireSupportedStyle(value.style().orElse(null), sheetName);
          default -> {}
        }
      }
    }
    return sheet.getCTWorksheet().getConditionalFormattingList().stream()
        .map(block -> (CTConditionalFormatting) block.copy())
        .toList();
  }

  private static void requireSupportedStyle(
      @org.jspecify.annotations.Nullable DifferentialStyleReport style, String sheetName) {
    if (style != null && !style.unsupportedFeatures().isEmpty()) {
      throw new IllegalArgumentException(
          "cannot copy sheet '"
              + sheetName
              + "': unsupported conditional-formatting differential-style features "
              + style.unsupportedFeatures());
    }
  }

  static void replaceConditionalFormatting(
      List<CTConditionalFormatting> blocks,
      ExcelSheet targetSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    var target = workbook.xssfWorkbook().getSheet(newSheetName);
    target.getCTWorksheet().setConditionalFormattingArray(new CTConditionalFormatting[0]);
    int index = workbook.xssfWorkbook().getSheetIndex(target);
    for (var original : blocks) {
      var block = target.getCTWorksheet().addNewConditionalFormatting();
      block.set(original);
      for (var rule : block.getCfRuleList()) {
        for (var threshold : thresholds(rule)) {
          if (threshold.getType()
                  == org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfvoType.FORMULA
              && threshold.getVal() != null) {
            threshold.setVal(
                ExcelSheetCopySupport.retargetFormula(
                    workbook,
                    threshold.getVal(),
                    FormulaType.CONDFORMAT,
                    index,
                    sourceSheetName,
                    newSheetName));
          }
        }
        for (int formula = 0; formula < rule.sizeOfFormulaArray(); formula++) {
          rule.setFormulaArray(
              formula,
              ExcelSheetCopySupport.retargetFormula(
                  workbook,
                  rule.getFormulaArray(formula),
                  FormulaType.CONDFORMAT,
                  index,
                  sourceSheetName,
                  newSheetName));
        }
      }
    }
  }

  private static List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfvo> thresholds(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule rule) {
    if (rule.isSetColorScale()) return rule.getColorScale().getCfvoList();
    if (rule.isSetDataBar()) return rule.getDataBar().getCfvoList();
    if (rule.isSetIconSet()) return rule.getIconSet().getCfvoList();
    return List.of();
  }
}
