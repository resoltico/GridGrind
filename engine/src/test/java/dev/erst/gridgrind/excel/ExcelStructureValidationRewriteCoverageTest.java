package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
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

      controller.insertRows(splitSheet, 2, 1);

      assertEquals(List.of("A2", "A4:A5"), dataValidationRanges(splitSheet));

      XSSFSheet formulaSheet = workbook.createSheet("FormulaRows");
      new ExcelDataValidationController()
          .setDataValidation(
              formulaSheet,
              "A2:A4",
              new ExcelDataValidationDefinition(
                  new ExcelDataValidationRule.CustomFormula("LEN(A2)>0"),
                  false,
                  false,
                  null,
                  null));
      setString(formulaSheet, "A2", "Ready");

      controller.insertRows(formulaSheet, 1, 1);

      ExcelDataValidationSnapshot.Supported supported =
          assertInstanceOf(
              ExcelDataValidationSnapshot.Supported.class,
              new ExcelDataValidationController()
                  .dataValidations(formulaSheet, new ExcelRangeSelection.All())
                  .getFirst());
      assertEquals(List.of("A3:A5"), supported.ranges());
      assertEquals(
          new ExcelDataValidationRule.CustomFormula("LEN(A3)>0"), supported.validation().rule());
    }
  }

  @Test
  void columnInsertSplitsValidationRangesAndRetargetsValidationFormulas() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet splitSheet = workbook.createSheet("SplitColumns");
      seedDataValidation(splitSheet);

      controller.insertColumns(splitSheet, 0, 1);

      assertEquals(List.of("B2:B4"), dataValidationRanges(splitSheet));

      XSSFSheet formulaSheet = workbook.createSheet("FormulaColumns");
      new ExcelDataValidationController()
          .setDataValidation(
              formulaSheet,
              "A1:A3",
              new ExcelDataValidationDefinition(
                  new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
                  false,
                  false,
                  null,
                  null));

      controller.insertColumns(formulaSheet, 0, 1);

      ExcelDataValidationSnapshot.Supported supported =
          assertInstanceOf(
              ExcelDataValidationSnapshot.Supported.class,
              new ExcelDataValidationController()
                  .dataValidations(formulaSheet, new ExcelRangeSelection.All())
                  .getFirst());
      assertEquals(List.of("B1:B3"), supported.ranges());
      assertEquals(
          new ExcelDataValidationRule.CustomFormula("LEN(B1)>0"), supported.validation().rule());
    }
  }
}
