package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for sheet drawing helper seams and validation. */
class ExcelSheetDrawingSupportTest {
  @Test
  void cleanupEmptyDrawingPatriarchIsANoOpForSheetsWithoutDrawings() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var support = new ExcelSheetDrawingSupport(sheet);

      assertDoesNotThrow(support::cleanupEmptyDrawingPatriarch);
      assertEquals(0, sheet.getRelations().size());
    }
  }

  @Test
  void drawingOperationsRejectBlankObjectNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var support = new ExcelSheetDrawingSupport(sheet);

      IllegalArgumentException failure =
          assertThrows(IllegalArgumentException.class, () -> support.drawingObjectPayload(" "));
      assertEquals("objectName must not be blank", failure.getMessage());
    }
  }
}
