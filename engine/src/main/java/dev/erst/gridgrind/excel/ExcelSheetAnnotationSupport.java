package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.ss.SpreadsheetVersion;
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

/** Hyperlink, comment, and cell-annotation support for one sheet facade. */
final class ExcelSheetAnnotationSupport {
  private final Sheet sheet;
  private final ExcelDrawingController drawingController;

  ExcelSheetAnnotationSupport(Sheet sheet, ExcelDrawingController drawingController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.drawingController =
        Objects.requireNonNull(drawingController, "drawingController must not be null");
  }

  void setHyperlink(String address, ExcelHyperlink hyperlink) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");

    CellReference cellReference = parseCellReference(address);
    Cell cell = getOrCreateCell(cellReference.getRow(), cellReference.getCol());
    cell.removeHyperlink();
    org.apache.poi.ss.usermodel.Hyperlink poiHyperlink =
        sheet.getWorkbook().getCreationHelper().createHyperlink(toPoi(hyperlink.type()));
    poiHyperlink.setAddress(toPoiTarget(hyperlink));
    cell.setHyperlink(poiHyperlink);
  }

  void clearHyperlink(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = optionalCell(address);
    if (cell != null) {
      cell.removeHyperlink();
    }
  }

  void setComment(String address, ExcelComment comment) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(comment, "comment must not be null");

    CellReference cellReference = parseCellReference(address);
    repairBrokenLegacyDrawingReference((XSSFSheet) sheet);
    Cell cell = getOrCreateCell(cellReference.getRow(), cellReference.getCol());
    cell.setCellComment(newComment(cellReference.getRow(), cellReference.getCol(), comment));
    ensureLegacyDrawingReference((XSSFSheet) sheet);
    drawingController.cleanupEmptyDrawingPatriarch((org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
  }

  void clearComment(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = optionalCell(address);
    if (cell != null) {
      clearCellComment(cell);
    }
    drawingController.cleanupEmptyDrawingPatriarch((org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
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

  ExcelCellMetadataSnapshot metadata(Cell cell) {
    return ExcelCellMetadataSnapshot.of(hyperlink(cell), commentSnapshot(cell));
  }

  List<WorkbookReadResult.CellHyperlink> hyperlinks(ExcelCellSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return switch (selection) {
      case ExcelCellSelection.AllUsedCells _ -> allUsedHyperlinks();
      case ExcelCellSelection.Selected selected -> selectedHyperlinks(selected.addresses());
    };
  }

  List<WorkbookReadResult.CellComment> comments(ExcelCellSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return switch (selection) {
      case ExcelCellSelection.AllUsedCells _ -> allUsedComments();
      case ExcelCellSelection.Selected selected -> selectedComments(selected.addresses());
    };
  }

  static ExcelHyperlink hyperlink(Cell cell) {
    return cell == null ? null : hyperlink(cell.getHyperlink());
  }

  static ExcelHyperlink hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    if (hyperlink == null || hyperlink.getType() == null) {
      return null;
    }
    return hyperlink(hyperlink.getType(), hyperlink.getAddress());
  }

  static ExcelHyperlink hyperlink(HyperlinkType hyperlinkType, String target) {
    if (hyperlinkType == null || target == null || target.isBlank()) {
      return null;
    }
    return switch (hyperlinkType) {
      case URL ->
          ExcelHyperlinkValidation.isValidUrlTarget(target) ? new ExcelHyperlink.Url(target) : null;
      case EMAIL ->
          ExcelHyperlinkValidation.isValidEmailTarget(target)
              ? new ExcelHyperlink.Email(target)
              : null;
      case FILE ->
          ExcelHyperlinkValidation.isValidFileTarget(target)
              ? new ExcelHyperlink.File(target)
              : null;
      case DOCUMENT -> new ExcelHyperlink.Document(target);
      case NONE -> null;
    };
  }

  static ExcelComment comment(Cell cell) {
    return cell == null ? null : comment(cell.getCellComment());
  }

  static ExcelComment comment(Comment comment) {
    if (comment == null || comment.getString() == null) {
      return null;
    }
    return comment(comment.getString().getString(), comment.getAuthor(), comment.isVisible());
  }

  static ExcelComment comment(String text, String author, boolean visible) {
    if (text == null || text.isBlank() || author == null || author.isBlank()) {
      return null;
    }
    return new ExcelComment(text, author, visible);
  }

  static ExcelCommentSnapshot commentSnapshot(Cell cell) {
    return cell == null
        ? null
        : commentSnapshot(cell.getSheet().getWorkbook(), cell.getCellComment());
  }

  static ExcelCommentSnapshot commentSnapshot(Comment comment) {
    return commentSnapshot(null, comment);
  }

  static HyperlinkType toPoi(ExcelHyperlinkType hyperlinkType) {
    return switch (hyperlinkType) {
      case URL -> HyperlinkType.URL;
      case EMAIL -> HyperlinkType.EMAIL;
      case FILE -> HyperlinkType.FILE;
      case DOCUMENT -> HyperlinkType.DOCUMENT;
    };
  }

  static String toPoiTarget(ExcelHyperlink hyperlink) {
    return switch (hyperlink) {
      case ExcelHyperlink.Url url -> url.target();
      case ExcelHyperlink.Email email -> "mailto:" + email.target();
      case ExcelHyperlink.File file -> ExcelFileHyperlinkTargets.toPoiAddress(file.path());
      case ExcelHyperlink.Document document -> document.target();
    };
  }

  private Comment newComment(int rowIndex, int columnIndex, ExcelComment comment) {
    ClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
    ExcelCommentAnchor authoredAnchor = comment.anchor();
    anchor.setRow1(authoredAnchor == null ? rowIndex : authoredAnchor.firstRow());
    anchor.setRow2(authoredAnchor == null ? rowIndex + 3 : authoredAnchor.lastRow());
    anchor.setCol1(authoredAnchor == null ? columnIndex : authoredAnchor.firstColumn());
    anchor.setCol2(authoredAnchor == null ? columnIndex + 3 : authoredAnchor.lastColumn());
    Comment poiComment = sheet.createDrawingPatriarch().createCellComment(anchor);
    poiComment.setAuthor(comment.author());
    poiComment.setVisible(comment.visible());
    poiComment.setString(
        comment.runs() == null
            ? new XSSFRichTextString(comment.text())
            : ExcelRichTextSupport.toPoiRichText(
                (XSSFWorkbook) sheet.getWorkbook(), comment.runs()));
    return poiComment;
  }

  private List<WorkbookReadResult.CellHyperlink> allUsedHyperlinks() {
    List<WorkbookReadResult.CellHyperlink> hyperlinks = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        ExcelHyperlink hyperlink = hyperlink(cell);
        if (hyperlink != null) {
          hyperlinks.add(
              new WorkbookReadResult.CellHyperlink(
                  new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                  hyperlink));
        }
      }
    }
    return List.copyOf(hyperlinks);
  }

  private List<WorkbookReadResult.CellHyperlink> selectedHyperlinks(List<String> addresses) {
    List<WorkbookReadResult.CellHyperlink> hyperlinks = new ArrayList<>();
    for (String address : addresses) {
      Cell cell = cellOrNull(address);
      if (cell == null) {
        continue;
      }
      ExcelHyperlink hyperlink = hyperlink(cell);
      if (hyperlink != null) {
        hyperlinks.add(new WorkbookReadResult.CellHyperlink(address, hyperlink));
      }
    }
    return List.copyOf(hyperlinks);
  }

  private List<WorkbookReadResult.CellComment> allUsedComments() {
    List<WorkbookReadResult.CellComment> comments = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        ExcelCommentSnapshot comment = commentSnapshot(cell);
        if (comment != null) {
          comments.add(
              new WorkbookReadResult.CellComment(
                  new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                  comment));
        }
      }
    }
    return List.copyOf(comments);
  }

  private List<WorkbookReadResult.CellComment> selectedComments(List<String> addresses) {
    List<WorkbookReadResult.CellComment> comments = new ArrayList<>();
    for (String address : addresses) {
      Cell cell = cellOrNull(address);
      if (cell == null) {
        continue;
      }
      ExcelCommentSnapshot comment = commentSnapshot(cell);
      if (comment != null) {
        comments.add(new WorkbookReadResult.CellComment(address, comment));
      }
    }
    return List.copyOf(comments);
  }

  private Cell cellOrNull(String address) {
    CellReference reference = parseCellReference(address);
    Row row = sheet.getRow(reference.getRow());
    return row == null ? null : row.getCell(reference.getCol());
  }

  private static ExcelCommentSnapshot commentSnapshot(Workbook workbook, Comment comment) {
    if (!(comment instanceof XSSFComment xssfComment)
        || xssfComment.getString() == null
        || xssfComment.getAuthor() == null
        || xssfComment.getAuthor().isBlank()) {
      return null;
    }
    ExcelComment plainComment =
        comment(
            xssfComment.getString().getString(), xssfComment.getAuthor(), xssfComment.isVisible());
    if (plainComment == null) {
      return null;
    }
    ExcelCommentAnchorSnapshot anchor = null;
    if (xssfComment.getClientAnchor() instanceof XSSFClientAnchor clientAnchor) {
      anchor =
          new ExcelCommentAnchorSnapshot(
              clientAnchor.getCol1(),
              clientAnchor.getRow1(),
              clientAnchor.getCol2(),
              clientAnchor.getRow2());
    }
    ExcelRichTextSnapshot runs =
        workbook == null
            ? null
            : ExcelRichTextSupport.snapshot(
                (XSSFWorkbook) workbook,
                xssfComment.getString(),
                WorkbookStyleRegistry.snapshotFont(((XSSFWorkbook) workbook).getFontAt(0)));
    return new ExcelCommentSnapshot(
        plainComment.text(), plainComment.author(), plainComment.visible(), runs, anchor);
  }

  private static void removeCommentFromTable(XSSFSheet sheet, CellAddress address) {
    for (POIXMLDocumentPart relation : sheet.getRelations()) {
      if (relation instanceof Comments comments) {
        comments.removeComment(address);
        return;
      }
    }
  }

  private static void removeCommentShapeIfPresent(XSSFSheet sheet, CellAddress address) {
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
    if (sheet.getCTWorksheet().isSetLegacyDrawing() && legacyDrawingRelationId(sheet) == null) {
      sheet.getCTWorksheet().unsetLegacyDrawing();
    }
  }

  static void ensureLegacyDrawingReference(XSSFSheet sheet) {
    String relationId = vmlDrawingRelationId(sheet);
    if (relationId == null) {
      return;
    }
    if (!sheet.getCTWorksheet().isSetLegacyDrawing()) {
      sheet.getCTWorksheet().addNewLegacyDrawing();
    }
    sheet.getCTWorksheet().getLegacyDrawing().setId(relationId);
  }

  static String legacyDrawingRelationId(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetLegacyDrawing()) {
      return null;
    }
    String legacyDrawingId = sheet.getCTWorksheet().getLegacyDrawing().getId();
    for (RelationPart relationPart : sheet.getRelationParts()) {
      if (relationPart.getDocumentPart() instanceof XSSFVMLDrawing
          && legacyDrawingId.equals(relationPart.getRelationship().getId())) {
        return legacyDrawingId;
      }
    }
    return null;
  }

  static String vmlDrawingRelationId(XSSFSheet sheet) {
    for (RelationPart relationPart : sheet.getRelationParts()) {
      if (relationPart.getDocumentPart() instanceof XSSFVMLDrawing) {
        return relationPart.getRelationship().getId();
      }
    }
    return null;
  }

  private CellReference parseCellReference(String address) {
    try {
      CellReference reference = new CellReference(address);
      requireValidCellReference(address, reference);
      return reference;
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  private static void requireValidCellReference(String address, CellReference cellReference) {
    int row = cellReference.getRow();
    int col = cellReference.getCol();
    if (row < 0
        || col < 0
        || row > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || col > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new InvalidCellAddressException(
          address, new IllegalArgumentException("not a valid A1-style cell address: " + address));
    }
  }

  private Cell optionalCell(String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      return null;
    }
    return row.getCell(cellReference.getCol());
  }

  private Cell getOrCreateCell(int rowIndex, int columnIndex) {
    return getOrCreateRow(rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }
}
