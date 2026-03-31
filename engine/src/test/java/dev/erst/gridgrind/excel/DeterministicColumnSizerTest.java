package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for deterministic column sizing behavior and failure handling. */
class DeterministicColumnSizerTest {
  @Test
  void contentWidthCharactersReturnsZeroForEmptyDisplayValues() {
    assertEquals(0.0d, DeterministicColumnSizer.contentWidthCharacters(""));
  }

  @Test
  void autoSizeSkipsZeroWidthDisplaysAndPropagatesNonPoiDisplayFailures() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Budget");
      sheet.createRow(0).createCell(0).setCellValue("");
      sheet.getRow(0).createCell(1).setCellFormula("1+1");

      FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      RuntimeException cause = new IllegalStateException("display failure");

      IllegalStateException exception =
          assertThrows(
              IllegalStateException.class,
              () ->
                  DeterministicColumnSizer.autoSize(
                      sheet,
                      "Budget",
                      new DataFormatter(),
                      FormulaRuntimeTestDouble.failingDisplay(evaluator, cause)));

      assertSame(cause, exception);
      assertEquals(2048, sheet.getColumnWidth(0));
    }
  }

  @Test
  void autoSizePropagatesNonPoiDisplayFailuresForNonFormulaCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Budget");
      sheet.createRow(0).createCell(0).setCellValue("Quarterly revenue");

      IllegalStateException cause = new IllegalStateException("formatter failure");

      IllegalStateException exception =
          assertThrows(
              IllegalStateException.class,
              () ->
                  DeterministicColumnSizer.autoSize(
                      sheet,
                      "Budget",
                      new DataFormatter() {
                        @Override
                        public String formatCellValue(Cell cell) {
                          throw cause;
                        }
                      },
                      FormulaRuntimeTestDouble.delegating(
                          workbook.getCreationHelper().createFormulaEvaluator())));

      assertSame(cause, exception);
    }
  }
}
