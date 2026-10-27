package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Comment mutation and snapshot helpers for one sheet. */
final class ExcelSheetCommentSupport {
  private final Sheet sheet;
  private final ExcelDrawingController drawingController;

  ExcelSheetCommentSupport(Sheet sheet, ExcelDrawingController drawingController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.drawingController =
        Objects.requireNonNull(drawingController, "drawingController must not be null");
  }

  void setComment(String address, ResolvedComment comment) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(comment, "comment must not be null");

    CellReference cellReference = ExcelSheetAddressSupport.parseCellReference(address);
    repairBrokenLegacyDrawingReference((XSSFSheet) sheet);
    Cell cell =
        ExcelSheetAddressSupport.getOrCreateCell(
            sheet, cellReference.getRow(), cellReference.getCol());
    cell.setCellComment(newComment(cellReference.getRow(), cellReference.getCol(), comment));
    ensureLegacyDrawingReference((XSSFSheet) sheet);
    drawingController.cleanupEmptyDrawingPatriarch((XSSFSheet) sheet);
  }

  void clearComment(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    ExcelSheetAddressSupport.optionalCell(sheet, address)
        .ifPresent(ExcelSheetCommentSupport::clearCellComment);
    drawingController.cleanupEmptyDrawingPatriarch((XSSFSheet) sheet);
  }

  static void clearCellComment(Cell cell) {
    Objects.requireNonNull(cell, "cell must not be null");
    if (!(cell.getSheet() instanceof XSSFSheet xssfSheet)
        || !(cell.getCellComment() instanceof XSSFComment)) {
      cell.removeCellComment();
      return;
    }

    CellAddress address = new CellAddress(cell);
    removeCommentFromTable(xssfSheet, address);
    removeCommentShapeIfPresent(xssfSheet, address);
  }

  static Optional<ResolvedComment> comment(Cell cell) {
    return cell == null ? Optional.empty() : comment(cell.getCellComment());
  }

  static Optional<ResolvedComment> comment(Comment comment) {
    return comment == null || comment.getString() == null
        ? Optional.empty()
        : comment(comment.getString().getString(), comment.getAuthor(), comment.isVisible());
  }

  static Optional<ResolvedComment> comment(String text, String author, boolean visible) {
    return text == null || text.isBlank() || author == null || author.isBlank()
        ? Optional.empty()
        : Optional.of(new ResolvedComment(text, author, visible));
  }

  static Optional<CommentReport> commentSnapshot(Cell cell) {
    return cell == null
        ? Optional.empty()
        : optionalCommentSnapshot(
            Optional.of(cell.getSheet().getWorkbook()), cell.getCellComment());
  }

  static Optional<CommentReport> commentSnapshot(Comment comment) {
    return optionalCommentSnapshot(Optional.empty(), comment);
  }

  static void removeCommentFromTable(XSSFSheet sheet, CellAddress address) {
    for (POIXMLDocumentPart relation : sheet.getRelations()) {
      if (relation instanceof Comments comments) {
        comments.removeComment(address);
        return;
      }
    }
  }

  static void removeCommentShapeIfPresent(XSSFSheet sheet, CellAddress address) {
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return;
    }
    var commentShape = vmlDrawing.findCommentShape(address.getRow(), address.getColumn());
    if (commentShape != null) {
      try (var cursor = commentShape.newCursor()) {
        cursor.removeXml();
      }
    }
  }

  static void repairBrokenLegacyDrawingReference(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().isSetLegacyDrawing() && legacyDrawingRelationId(sheet).isEmpty()) {
      sheet.getCTWorksheet().unsetLegacyDrawing();
    }
  }

  static void ensureLegacyDrawingReference(XSSFSheet sheet) {
    String relationId = vmlDrawingRelationId(sheet).orElse(null);
    if (relationId == null) {
      return;
    }
    if (!sheet.getCTWorksheet().isSetLegacyDrawing()) {
      sheet.getCTWorksheet().addNewLegacyDrawing();
    }
    sheet.getCTWorksheet().getLegacyDrawing().setId(relationId);
  }

  static Optional<String> legacyDrawingRelationId(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetLegacyDrawing()) {
      return Optional.empty();
    }
    String legacyDrawingId = sheet.getCTWorksheet().getLegacyDrawing().getId();
    for (RelationPart relationPart : sheet.getRelationParts()) {
      if (relationPart.getDocumentPart() instanceof XSSFVMLDrawing
          && legacyDrawingId.equals(relationPart.getRelationship().getId())) {
        return Optional.of(legacyDrawingId);
      }
    }
    return Optional.empty();
  }

  static Optional<String> vmlDrawingRelationId(XSSFSheet sheet) {
    for (RelationPart relationPart : sheet.getRelationParts()) {
      if (relationPart.getDocumentPart() instanceof XSSFVMLDrawing) {
        return Optional.of(relationPart.getRelationship().getId());
      }
    }
    return Optional.empty();
  }

  private Comment newComment(int rowIndex, int columnIndex, ResolvedComment comment) {
    ExcelCellTextLimits.requireSupportedLength(comment.text(), "comment.text"); // LIM-010
    ClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
    Optional<CommentAnchorSpec> authoredAnchor = comment.anchor();
    anchor.setRow1(authoredAnchor.map(CommentAnchorSpec::firstRow).orElse(rowIndex));
    anchor.setRow2(authoredAnchor.map(CommentAnchorSpec::lastRow).orElse(rowIndex + 3));
    anchor.setCol1(authoredAnchor.map(CommentAnchorSpec::firstColumn).orElse(columnIndex));
    anchor.setCol2(authoredAnchor.map(CommentAnchorSpec::lastColumn).orElse(columnIndex + 3));
    Comment poiComment = sheet.createDrawingPatriarch().createCellComment(anchor);
    poiComment.setAuthor(comment.author());
    poiComment.setVisible(comment.visible());
    poiComment.setString(
        comment.runs().isEmpty()
            ? new XSSFRichTextString(comment.text())
            : ExcelRichTextSupport.toPoiRichText(
                (XSSFWorkbook) sheet.getWorkbook(), comment.runs().orElseThrow()));
    return poiComment;
  }

  private List<CellCommentReport> allUsedComments() {
    List<CellCommentReport> comments = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        Optional<CommentReport> comment = commentSnapshot(cell);
        if (comment.isPresent()) {
          comments.add(
              new CellCommentReport(
                  new dev.erst.gridgrind.core.vocabulary.CellAddress(
                      new CellReference(cell.getRowIndex(), cell.getColumnIndex())
                          .formatAsString()),
                  comment.orElseThrow()));
        }
      }
    }
    return List.copyOf(comments);
  }

  List<CellCommentReport> comments() {
    return allUsedComments();
  }

  List<CellCommentReport> comments(List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    return selectedComments(List.copyOf(addresses));
  }

  private List<CellCommentReport> selectedComments(
      List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    List<CellCommentReport> comments = new ArrayList<>();
    for (dev.erst.gridgrind.core.vocabulary.CellAddress address : addresses) {
      Cell cell = ExcelSheetAddressSupport.cellOrNull(sheet, address.value()).orElse(null);
      if (cell == null) {
        continue;
      }
      Optional<CommentReport> comment = commentSnapshot(cell);
      if (comment.isPresent()) {
        comments.add(new CellCommentReport(address, comment.orElseThrow()));
      }
    }
    return List.copyOf(comments);
  }

  private static Optional<CommentReport> optionalCommentSnapshot(
      Optional<Workbook> workbook, Comment comment) {
    if (!(comment instanceof XSSFComment xssfComment)
        || xssfComment.getString() == null
        || xssfComment.getAuthor() == null
        || xssfComment.getAuthor().isBlank()) {
      return Optional.empty();
    }
    Optional<ResolvedComment> plainComment =
        comment(
            xssfComment.getString().getString(), xssfComment.getAuthor(), xssfComment.isVisible());
    if (plainComment.isEmpty()) {
      return Optional.empty();
    }
    Optional<CommentAnchorReport> anchor = Optional.empty();
    if (xssfComment.getClientAnchor() instanceof XSSFClientAnchor clientAnchor) {
      anchor =
          Optional.of(
              new CommentAnchorReport(
                  clientAnchor.getCol1(),
                  clientAnchor.getRow1(),
                  clientAnchor.getCol2(),
                  clientAnchor.getRow2()));
    }
    Optional<List<RichTextRunReport>> runs =
        workbook.isEmpty()
            ? Optional.empty()
            : ExcelRichTextSupport.snapshot(
                (XSSFWorkbook) workbook.orElseThrow(),
                xssfComment.getString(),
                ExcelCellStyleSnapshotSupport.snapshotFont(
                    (org.apache.poi.xssf.usermodel.XSSFFont) workbook.get().getFontAt(0)));
    return Optional.of(
        new CommentReport(
            plainComment.orElseThrow().text(),
            plainComment.orElseThrow().author(),
            plainComment.orElseThrow().visible(),
            runs,
            anchor));
  }
}
