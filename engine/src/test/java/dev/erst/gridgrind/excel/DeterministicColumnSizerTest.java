package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
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
  void autoSizeSkipsZeroWidthDisplaysWithoutEvaluatingFormulaCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Budget");
      sheet.createRow(0).createCell(0).setCellValue("");
      sheet.getRow(0).createCell(1).setCellFormula("1+1");

      DeterministicColumnSizer.autoSize(sheet, new DataFormatter());
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
                      new DataFormatter() {
                        @Override
                        public String formatCellValue(Cell cell) {
                          throw cause;
                        }
                      }));

      assertSame(cause, exception);
    }
  }

  @Test
  void autoSizeDoesNotPersistFormulaEvaluationOutsideCalculationPolicy() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Budget");
      sheet.createRow(0).createCell(0).setCellValue(1.0d);
      sheet.getRow(0).createCell(1).setCellFormula("A1+5");

      DeterministicColumnSizer.autoSize(sheet, new DataFormatter());

      assertFalse(worksheetXml(workbook).contains("<v>6.0</v>"));
    }
  }

  private static String worksheetXml(XSSFWorkbook workbook) throws java.io.IOException {
    ByteArrayOutputStream bytes = new ByteArrayOutputStream();
    workbook.write(bytes);
    try (ZipInputStream zip = new ZipInputStream(new ByteArrayInputStream(bytes.toByteArray()))) {
      ZipEntry entry = zip.getNextEntry();
      while (entry != null) {
        if ("xl/worksheets/sheet1.xml".equals(entry.getName())) {
          return new String(zip.readAllBytes(), StandardCharsets.UTF_8);
        }
        entry = zip.getNextEntry();
      }
    }
    throw new AssertionError("worksheet XML was not written");
  }
}
