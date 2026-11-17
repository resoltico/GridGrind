package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

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
  void annotationMutationsUseCanonicalCellCoordinates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var comments = new CommentAdapter(sheet);
      var links = new LinkAdapter(sheet);
      var address = new dev.erst.gridgrind.core.vocabulary.CellAddress(7, 27);
      links.setHyperlink(address, new HyperlinkTarget.Url("https://example.com"));
      comments.setComment(address, new ResolvedComment("Review", "GridGrind", false));
      var nativeCell = sheet.getRow(7).getCell(27);
      assertEquals("https://example.com", nativeCell.getHyperlink().getAddress());
      assertEquals("Review", nativeCell.getCellComment().getString().getString());
      var alias = new dev.erst.gridgrind.core.vocabulary.CellAddress("$ab$8");
      links.clearHyperlink(alias);
      comments.clearComment(alias);
      assertNull(nativeCell.getHyperlink());
      assertNull(nativeCell.getCellComment());
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
