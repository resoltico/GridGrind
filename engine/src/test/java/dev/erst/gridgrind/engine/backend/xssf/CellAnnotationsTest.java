package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.links.LinkAdapter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for sheet-annotation helper validation and address handling. */
class CellAnnotationsTest {
  @Test
  void annotationMutationsRejectOutOfBoundsCellAddresses() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var comments = new CommentAdapter(sheet);
      var links = new LinkAdapter(sheet);

      InvalidCellAddressException zeroRowFailure =
          assertThrows(InvalidCellAddressException.class, () -> links.clearHyperlink("A0"));
      assertEquals("A0", zeroRowFailure.address());

      InvalidCellAddressException missingColumnFailure =
          assertThrows(InvalidCellAddressException.class, () -> comments.clearComment("1"));
      assertEquals("1", missingColumnFailure.address());

      InvalidCellAddressException hyperlinkFailure =
          assertThrows(
              InvalidCellAddressException.class,
              () -> links.setHyperlink("A1048577", new HyperlinkTarget.Url("https://example.com")));
      assertEquals("A1048577", hyperlinkFailure.address());

      InvalidCellAddressException commentFailure =
          assertThrows(
              InvalidCellAddressException.class,
              () -> comments.setComment("XFE1", new ResolvedComment("Review", "GridGrind", false)));
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

      CommentAdapter.clearCellComment(cell);

      assertNull(cell.getCellComment());
    }
  }

  @Test
  void removeCommentShapeIfPresentReturnsWhenSheetHasNoVmlDrawing() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      assertNull(sheet.getVMLDrawing(false));
      assertDoesNotThrow(
          () -> CommentAdapter.removeCommentShapeIfPresent(sheet, new CellAddress("A1")));
      assertNull(sheet.getVMLDrawing(false));
    }
  }
}
