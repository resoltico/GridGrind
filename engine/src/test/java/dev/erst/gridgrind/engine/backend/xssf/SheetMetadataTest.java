package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for sheet metadata helper construction and validation. */
class SheetMetadataTest {
  @Test
  void defaultConstructorProvidesWorkingControllers() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var support = new SheetMetadata(workbook.createSheet("Ops"));

      assertEquals(0, support.conditionalFormattingBlockCount());
    }
  }
}
