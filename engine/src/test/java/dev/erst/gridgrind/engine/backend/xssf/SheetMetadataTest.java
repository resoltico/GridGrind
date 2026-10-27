package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import java.util.List;
import java.util.Optional;
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

  @Test
  void metadataMutationsRejectBlankRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var support = new SheetMetadata(workbook.createSheet("Ops"));

      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  support.setDataValidation(
                      " ",
                      new ResolvedDataValidation(
                          new DataValidationRuleSpec.ExplicitList(List.of("Open")),
                          false,
                          false,
                          Optional.empty(),
                          Optional.empty())));
      assertEquals("range must not be blank", validationFailure.getMessage());

      IllegalArgumentException autofilterFailure =
          assertThrows(IllegalArgumentException.class, () -> support.setAutofilter(" "));
      assertEquals("range must not be blank", autofilterFailure.getMessage());
    }
  }
}
