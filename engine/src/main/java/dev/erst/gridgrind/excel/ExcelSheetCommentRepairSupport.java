package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.NavigableMap;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

/** Repairs authoritative sheet comment state after POI mutations that can corrupt OOXML. */
final class ExcelSheetCommentRepairSupport {
  private static final Comparator<CellAddress> CELL_ADDRESS_ORDER =
      Comparator.comparingInt(CellAddress::getRow).thenComparingInt(CellAddress::getColumn);

  private final XSSFSheet sheet;
  private final ExcelSheetAnnotationSupport annotationSupport;
  private final ExcelDrawingController drawingController;

  ExcelSheetCommentRepairSupport(XSSFSheet sheet) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    drawingController = new ExcelDrawingController();
    annotationSupport = new ExcelSheetAnnotationSupport(sheet, drawingController);
  }

  boolean hasPersistedComments() {
    return !sheet.getCellComments().isEmpty()
        || !rawCommentAddresses(commentsTable(sheet).orElse(null)).isEmpty();
  }

  List<CommentRewriteSnapshot> expectedCommentsAfterInsertColumns(
      int columnIndex, int columnCount) {
    return rewriteCommentsAfterInsertColumns(
        snapshotCommentsForRewrite(), columnIndex, columnCount);
  }

  List<CommentRewriteSnapshot> expectedCommentsAfterDeleteColumns(ExcelColumnSpan columns) {
    return rewriteCommentsAfterDeleteColumns(snapshotCommentsForRewrite(), columns);
  }

  List<CommentRewriteSnapshot> expectedCommentsAfterShiftColumns(
      ExcelColumnSpan columns, int delta) {
    return rewriteCommentsAfterShiftColumns(snapshotCommentsForRewrite(), columns, delta);
  }

  void replaceComments(List<CommentRewriteSnapshot> expectedComments) {
    Objects.requireNonNull(expectedComments, "expectedComments must not be null");
    clearAllPersistedComments();
    for (CommentRewriteSnapshot expectedComment : expectedComments) {
      Objects.requireNonNull(expectedComment, "expectedComments must not contain nulls");
      replaceOneComment(expectedComment);
    }
  }

  static List<CommentRewriteSnapshot> commentRewriteSnapshots(
      List<WorkbookSheetResult.CellComment> comments) {
    Objects.requireNonNull(comments, "comments must not be null");
    List<CommentRewriteSnapshot> snapshots = new ArrayList<>(comments.size());
    for (WorkbookSheetResult.CellComment comment : comments) {
      Objects.requireNonNull(comment, "comments must not contain nulls");
      snapshots.add(CommentRewriteSnapshot.from(comment));
    }
    return List.copyOf(snapshots);
  }

  static List<WorkbookSheetResult.CellComment> commentsAfterInsertColumns(
      List<WorkbookSheetResult.CellComment> comments, int columnIndex, int columnCount) {
    return toReadComments(
        rewriteCommentsAfterInsertColumns(
            commentRewriteSnapshots(comments), columnIndex, columnCount));
  }

  static List<WorkbookSheetResult.CellComment> commentsAfterDeleteColumns(
      List<WorkbookSheetResult.CellComment> comments, ExcelColumnSpan columns) {
    return toReadComments(
        rewriteCommentsAfterDeleteColumns(commentRewriteSnapshots(comments), columns));
  }

  static List<WorkbookSheetResult.CellComment> commentsAfterShiftColumns(
      List<WorkbookSheetResult.CellComment> comments, ExcelColumnSpan columns, int delta) {
    return toReadComments(
        rewriteCommentsAfterShiftColumns(commentRewriteSnapshots(comments), columns, delta));
  }

  private void replaceOneComment(CommentRewriteSnapshot expectedComment) {
    if (expectedComment.isAuthoringCompatible()) {
      annotationSupport.setComment(expectedComment.address(), expectedComment.toAuthoringComment());
      return;
    }
    setRawComment(expectedComment);
  }

  private void setRawComment(CommentRewriteSnapshot expectedComment) {
    CellReference reference = new CellReference(expectedComment.address());
    ExcelSheetAnnotationSupport.repairBrokenLegacyDrawingReference(sheet);
    Cell cell = getOrCreateCell(sheet, reference.getRow(), reference.getCol());
    cell.setCellComment(newRawComment(reference.getRow(), reference.getCol(), expectedComment));
    ExcelSheetAnnotationSupport.ensureLegacyDrawingReference(sheet);
    drawingController.cleanupEmptyDrawingPatriarch(sheet);
  }

  private XSSFComment newRawComment(
      int rowIndex, int columnIndex, CommentRewriteSnapshot expectedComment) {
    ClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
    ExcelCommentAnchorSnapshot authoredAnchor = expectedComment.anchor();
    anchor.setRow1(authoredAnchor == null ? rowIndex : authoredAnchor.firstRow());
    anchor.setRow2(authoredAnchor == null ? rowIndex + 3 : authoredAnchor.lastRow());
    anchor.setCol1(authoredAnchor == null ? columnIndex : authoredAnchor.firstColumn());
    anchor.setCol2(authoredAnchor == null ? columnIndex + 3 : authoredAnchor.lastColumn());
    XSSFComment poiComment = (XSSFComment) sheet.createDrawingPatriarch().createCellComment(anchor);
    poiComment.setAuthor(expectedComment.author());
    poiComment.setVisible(expectedComment.visible());
    poiComment.setString(
        expectedComment.runs() == null
            ? new XSSFRichTextString(expectedComment.text())
            : ExcelRichTextSupport.toPoiRichText(
                (XSSFWorkbook) sheet.getWorkbook(), toAuthoringRichText(expectedComment.runs())));
    return poiComment;
  }

  private List<CommentRewriteSnapshot> snapshotCommentsForRewrite() {
    NavigableMap<CellAddress, XSSFComment> commentsByAddress = new TreeMap<>(CELL_ADDRESS_ORDER);
    commentsByAddress.putAll(sheet.getCellComments());
    List<CommentRewriteSnapshot> comments = new ArrayList<>(commentsByAddress.size());
    for (var entry : commentsByAddress.entrySet()) {
      comments.add(
          CommentRewriteSnapshot.from(
              entry.getKey(), entry.getValue(), (XSSFWorkbook) sheet.getWorkbook()));
    }
    return List.copyOf(comments);
  }

  private void clearAllPersistedComments() {
    CommentsTable commentsTable = commentsTable(sheet).orElse(null);
    List<CellAddress> rawAddresses = rawCommentAddresses(commentsTable);
    if (rawAddresses.isEmpty()) {
      return;
    }

    Set<CellAddress> uniqueAddresses = new LinkedHashSet<>(rawAddresses);
    for (CellAddress address : uniqueAddresses) {
      while (commentsTable.removeComment(address)) {
        // Continue until every duplicate CTComment for the same ref has been removed.
      }
    }

    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return;
    }
    for (CellAddress address : uniqueAddresses) {
      removeAllCommentShapes(vmlDrawing, address);
    }
  }

  private static List<CommentRewriteSnapshot> rewriteCommentsAfterInsertColumns(
      List<CommentRewriteSnapshot> comments, int columnIndex, int columnCount) {
    Objects.requireNonNull(comments, "comments must not be null");

    NavigableMap<CellAddress, CommentRewriteSnapshot> shifted = new TreeMap<>(CELL_ADDRESS_ORDER);
    for (CommentRewriteSnapshot comment : comments) {
      Objects.requireNonNull(comment, "comments must not contain nulls");
      CellAddress address = parsedAddress(comment.address());
      int targetColumn =
          address.getColumn() >= columnIndex
              ? address.getColumn() + columnCount
              : address.getColumn();
      shifted.put(at(address.getRow(), targetColumn), comment.at(address.getRow(), targetColumn));
    }
    return List.copyOf(shifted.values());
  }

  private static List<CommentRewriteSnapshot> rewriteCommentsAfterDeleteColumns(
      List<CommentRewriteSnapshot> comments, ExcelColumnSpan columns) {
    Objects.requireNonNull(comments, "comments must not be null");
    Objects.requireNonNull(columns, "columns must not be null");

    NavigableMap<CellAddress, CommentRewriteSnapshot> shifted = new TreeMap<>(CELL_ADDRESS_ORDER);
    for (CommentRewriteSnapshot comment : comments) {
      Objects.requireNonNull(comment, "comments must not contain nulls");
      CellAddress address = parsedAddress(comment.address());
      int column = address.getColumn();
      if (column >= columns.firstColumnIndex() && column <= columns.lastColumnIndex()) {
        continue;
      }
      int targetColumn = column > columns.lastColumnIndex() ? column - columns.count() : column;
      shifted.put(at(address.getRow(), targetColumn), comment.at(address.getRow(), targetColumn));
    }
    return List.copyOf(shifted.values());
  }

  private static List<CommentRewriteSnapshot> rewriteCommentsAfterShiftColumns(
      List<CommentRewriteSnapshot> comments, ExcelColumnSpan columns, int delta) {
    Objects.requireNonNull(comments, "comments must not be null");
    Objects.requireNonNull(columns, "columns must not be null");
    if (delta == 0) {
      return List.copyOf(comments);
    }

    NavigableMap<CellAddress, CommentRewriteSnapshot> shifted = new TreeMap<>(CELL_ADDRESS_ORDER);
    for (CommentRewriteSnapshot comment : comments) {
      Objects.requireNonNull(comment, "comments must not contain nulls");
      CellAddress address = parsedAddress(comment.address());
      if (inSourceColumns(address.getColumn(), columns)) {
        continue;
      }
      if (inOverwrittenColumns(address.getColumn(), columns, delta)) {
        continue;
      }
      shifted.put(address, comment);
    }
    for (CommentRewriteSnapshot comment : comments) {
      CellAddress address = parsedAddress(comment.address());
      if (!inSourceColumns(address.getColumn(), columns)) {
        continue;
      }
      shifted.put(
          at(address.getRow(), address.getColumn() + delta),
          comment.at(address.getRow(), address.getColumn() + delta));
    }
    return List.copyOf(shifted.values());
  }

  private static List<WorkbookSheetResult.CellComment> toReadComments(
      List<CommentRewriteSnapshot> comments) {
    List<WorkbookSheetResult.CellComment> readComments = new ArrayList<>(comments.size());
    for (CommentRewriteSnapshot comment : comments) {
      readComments.add(comment.toReadComment());
    }
    return List.copyOf(readComments);
  }

  private static Optional<CommentsTable> commentsTable(XSSFSheet sheet) {
    for (POIXMLDocumentPart relation : sheet.getRelations()) {
      if (relation instanceof CommentsTable commentsTable) {
        return Optional.of(commentsTable);
      }
    }
    return Optional.empty();
  }

  static List<CellAddress> rawCommentAddresses(CommentsTable commentsTable) {
    if (commentsTable == null || commentsTable.getCTComments().getCommentList() == null) {
      return List.of();
    }
    List<CellAddress> addresses = new ArrayList<>();
    for (CTComment comment : commentsTable.getCTComments().getCommentList().getCommentArray()) {
      addresses.add(new CellAddress(comment.getRef()));
    }
    return List.copyOf(addresses);
  }

  private static void removeAllCommentShapes(XSSFVMLDrawing vmlDrawing, CellAddress address) {
    var commentShape = vmlDrawing.findCommentShape(address.getRow(), address.getColumn());
    while (commentShape != null) {
      try (var cursor = commentShape.newCursor()) {
        cursor.removeXml();
      }
      commentShape = vmlDrawing.findCommentShape(address.getRow(), address.getColumn());
    }
  }

  private static CellAddress parsedAddress(String address) {
    return new CellAddress(address);
  }

  private static CellAddress at(int rowIndex, int columnIndex) {
    return new CellAddress(rowIndex, columnIndex);
  }

  private static String addressString(int rowIndex, int columnIndex) {
    return new CellReference(rowIndex, columnIndex).formatAsString();
  }

  private static boolean inSourceColumns(int column, ExcelColumnSpan columns) {
    return column >= columns.firstColumnIndex() && column <= columns.lastColumnIndex();
  }

  private static boolean inOverwrittenColumns(int column, ExcelColumnSpan columns, int delta) {
    if (delta > 0) {
      return column > columns.lastColumnIndex() && column <= columns.lastColumnIndex() + delta;
    }
    return column >= columns.firstColumnIndex() + delta && column < columns.firstColumnIndex();
  }

  private static Cell getOrCreateCell(XSSFSheet sheet, int rowIndex, int columnIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    Cell cell = row.getCell(columnIndex);
    return cell == null ? row.createCell(columnIndex) : cell;
  }

  private static ExcelRichText toAuthoringRichText(ExcelRichTextSnapshot runs) {
    return new ExcelRichText(
        runs.runs().stream().map(ExcelSheetCommentRepairSupport::toAuthoringRun).toList());
  }

  private static ExcelRichTextRun toAuthoringRun(ExcelRichTextRunSnapshot run) {
    return new ExcelRichTextRun(run.text(), toAuthoringFont(run.font()));
  }

  private static ExcelCellFont toAuthoringFont(ExcelCellFontSnapshot font) {
    return new ExcelCellFont(
        font.bold(),
        font.italic(),
        font.fontName(),
        font.fontHeight(),
        ExcelColorSupport.copyOf(font.fontColor()),
        font.underline(),
        font.strikeout());
  }

  record CommentRewriteSnapshot(
      String address,
      String text,
      String author,
      boolean visible,
      ExcelRichTextSnapshot runs,
      ExcelCommentAnchorSnapshot anchor) {
    CommentRewriteSnapshot {
      Objects.requireNonNull(address, "address must not be null");
      text = text == null ? "" : text;
      author = author == null ? "" : author;
    }

    static CommentRewriteSnapshot from(WorkbookSheetResult.CellComment comment) {
      return new CommentRewriteSnapshot(
          comment.address(),
          comment.comment().text(),
          comment.comment().author(),
          comment.comment().visible(),
          comment.comment().runs(),
          comment.comment().anchor());
    }

    static CommentRewriteSnapshot from(
        CellAddress address, XSSFComment comment, XSSFWorkbook workbook) {
      ExcelCommentAnchorSnapshot anchor = null;
      if (comment.getClientAnchor() instanceof XSSFClientAnchor clientAnchor) {
        anchor =
            new ExcelCommentAnchorSnapshot(
                clientAnchor.getCol1(),
                clientAnchor.getRow1(),
                clientAnchor.getCol2(),
                clientAnchor.getRow2());
      }
      XSSFRichTextString richText = comment.getString();
      ExcelRichTextSnapshot runs =
          richText == null
              ? null
              : ExcelRichTextSupport.snapshot(
                  workbook, richText, WorkbookStyleRegistry.snapshotFont(workbook.getFontAt(0)));
      return new CommentRewriteSnapshot(
          addressString(address.getRow(), address.getColumn()),
          richText == null ? "" : richText.getString(),
          comment.getAuthor(),
          comment.isVisible(),
          runs,
          anchor);
    }

    CommentRewriteSnapshot at(int rowIndex, int columnIndex) {
      return new CommentRewriteSnapshot(
          addressString(rowIndex, columnIndex), text, author, visible, runs, anchor);
    }

    boolean isAuthoringCompatible() {
      return !text.isBlank() && !author.isBlank();
    }

    ExcelComment toAuthoringComment() {
      return new ExcelComment(
          text,
          author,
          visible,
          runs == null ? null : toAuthoringRichText(runs),
          anchor == null
              ? null
              : new ExcelCommentAnchor(
                  anchor.firstColumn(), anchor.firstRow(), anchor.lastColumn(), anchor.lastRow()));
    }

    WorkbookSheetResult.CellComment toReadComment() {
      return new WorkbookSheetResult.CellComment(
          address, new ExcelCommentSnapshot(text, author, visible, runs, anchor));
    }
  }
}
