package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationController;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Validation-range rewrite coverage for row and column insertion. */
class ExcelStructureValidationRewriteCoverageTest extends ExcelRowColumnStructureTestSupport {
  @Test
  void rowInsertSplitsValidationRangesAndRetargetsValidationFormulas() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet splitSheet = workbook.createSheet("SplitRows");
      seedDataValidation(splitSheet);

      rowController.insertRows(splitSheet, 2, 1);

      assertEquals(List.of("A2", "A4:A5"), dataValidationRanges(splitSheet));

      XSSFSheet formulaSheet = workbook.createSheet("FormulaRows");
      new ExcelDataValidationController()
          .setDataValidation(
              formulaSheet,
              "A2:A4",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.CustomFormula(
                      ExcelComparisonFormulaSupport.normalizeFormula("LEN(A2)>0", "formula")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));
      setString(formulaSheet, "A2", "Ready");

      rowController.insertRows(formulaSheet, 1, 1);

      DataValidationEntryReport.Supported supported =
          assertInstanceOf(
              DataValidationEntryReport.Supported.class,
              new ExcelDataValidationController().dataValidations(formulaSheet).getFirst());
      assertEquals(
          List.of("A3:A5"),
          supported.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
      assertEquals(
          new DataValidationRuleSpec.CustomFormula(
              ExcelComparisonFormulaSupport.normalizeFormula("LEN(A3)>0", "formula")),
          supported.validation().rule());
    }
  }

  @Test
  void columnInsertSplitsValidationRangesAndRetargetsValidationFormulas() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet splitSheet = workbook.createSheet("SplitColumns");
      seedDataValidation(splitSheet);

      columnController.insertColumns(splitSheet, 0, 1);

      assertEquals(List.of("B2:B4"), dataValidationRanges(splitSheet));

      XSSFSheet formulaSheet = workbook.createSheet("FormulaColumns");
      new ExcelDataValidationController()
          .setDataValidation(
              formulaSheet,
              "A1:A3",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.CustomFormula(
                      ExcelComparisonFormulaSupport.normalizeFormula("LEN(A1)>0", "formula")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));

      columnController.insertColumns(formulaSheet, 0, 1);

      DataValidationEntryReport.Supported supported =
          assertInstanceOf(
              DataValidationEntryReport.Supported.class,
              new ExcelDataValidationController().dataValidations(formulaSheet).getFirst());
      assertEquals(
          List.of("B1:B3"),
          supported.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
      assertEquals(
          new DataValidationRuleSpec.CustomFormula(
              ExcelComparisonFormulaSupport.normalizeFormula("LEN(B1)>0", "formula")),
          supported.validation().rule());
    }
  }
}
