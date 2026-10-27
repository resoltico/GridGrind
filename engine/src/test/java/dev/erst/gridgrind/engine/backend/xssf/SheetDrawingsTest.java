package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Bound drawing operations preserve their object-name validation. */
class SheetDrawingsTest {
  @Test
  void drawingOperationsRejectBlankObjectNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var drawings = new SheetDrawings(sheet);
      var failure =
          assertThrows(IllegalArgumentException.class, () -> drawings.drawingObjectPayload(" "));
      assertEquals("objectName must not be blank", failure.getMessage());
    }
  }
}
