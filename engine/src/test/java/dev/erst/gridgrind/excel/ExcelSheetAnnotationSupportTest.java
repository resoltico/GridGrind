package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
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

  @Test
  void clearCellCommentHandlesAnchorlessCommentsWithoutVmlDrawing() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var cell = sheet.createRow(0).createCell(0);
      var commentsTable = new CommentsTable();
      commentsTable.setSheet(sheet);

      XSSFComment anchorless =
          new XSSFComment(commentsTable, commentsTable.newComment(new CellAddress("A1")), null);
      anchorless.setString("Review");
      anchorless.setAuthor("GridGrind");
      anchorless.setVisible(true);
      cell.setCellComment(anchorless);

      ExcelSheetAnnotationSupport.clearCellComment(cell);

      assertNull(cell.getCellComment());
    }
  }

  @Test
  void removeCommentShapeIfPresentReturnsWhenSheetHasNoVmlDrawing() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      assertNull(sheet.getVMLDrawing(false));
      assertDoesNotThrow(
          () ->
              ExcelSheetAnnotationSupport.removeCommentShapeIfPresent(
                  sheet, new CellAddress("A1")));
      assertNull(sheet.getVMLDrawing(false));
    }
  }
}
