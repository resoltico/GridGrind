package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for sheet-annotation helper validation and address handling. */
class ExcelSheetAnnotationSupportTest {
  @Test
  void annotationMutationsRejectOutOfBoundsCellAddresses() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var support = new ExcelSheetAnnotationSupport(sheet, new ExcelDrawingController());

      InvalidCellAddressException zeroRowFailure =
          assertThrows(InvalidCellAddressException.class, () -> support.clearHyperlink("A0"));
      assertEquals("A0", zeroRowFailure.address());

      InvalidCellAddressException missingColumnFailure =
          assertThrows(InvalidCellAddressException.class, () -> support.clearComment("1"));
      assertEquals("1", missingColumnFailure.address());

      InvalidCellAddressException hyperlinkFailure =
          assertThrows(
              InvalidCellAddressException.class,
              () ->
                  support.setHyperlink("A1048577", new ExcelHyperlink.Url("https://example.com")));
      assertEquals("A1048577", hyperlinkFailure.address());

      InvalidCellAddressException commentFailure =
          assertThrows(
              InvalidCellAddressException.class,
              () -> support.setComment("XFE1", new ExcelComment("Review", "GridGrind", false)));
      assertEquals("XFE1", commentFailure.address());
    }
  }
}
