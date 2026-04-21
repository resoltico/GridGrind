package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for sheet metadata helper construction and validation. */
class ExcelSheetMetadataSupportTest {
  @Test
  void defaultConstructorProvidesWorkingControllers() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var support = new ExcelSheetMetadataSupport(workbook.createSheet("Ops"));

      assertEquals(0, support.conditionalFormattingBlockCount());
    }
  }

  @Test
  void metadataMutationsRejectBlankRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var support = new ExcelSheetMetadataSupport(workbook.createSheet("Ops"));

      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  support.setDataValidation(
                      " ",
                      new ExcelDataValidationDefinition(
                          new ExcelDataValidationRule.ExplicitList(List.of("Open")),
                          false,
                          false,
                          null,
                          null),
                      null));
      assertEquals("range must not be blank", validationFailure.getMessage());

      IllegalArgumentException autofilterFailure =
          assertThrows(IllegalArgumentException.class, () -> support.setAutofilter(" ", null));
      assertEquals("range must not be blank", autofilterFailure.getMessage());
    }
  }
}
