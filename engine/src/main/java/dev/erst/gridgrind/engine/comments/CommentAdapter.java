package dev.erst.gridgrind.engine.comments;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.engine.cells.CellAccess;
import dev.erst.gridgrind.engine.cells.CellTextLimits;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
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
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

/** Owns comment read/write, VML relationships and before-mutation structural repair state. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL
})
public final class CommentAdapter {
  private final XSSFSheet sheet;

  public CommentAdapter(XSSFSheet sheet) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
  }

  public void setComment(String address, ResolvedComment comment) {
    CellAccess.requireAddress(address);
    Objects.requireNonNull(comment, "comment must not be null");

    CellReference cellReference = CellAccess.parseCellReference(address);
    repairBrokenLegacyDrawingReference((XSSFSheet) sheet);
    Cell cell = CellAccess.getOrCreateCell(sheet, cellReference.getRow(), cellReference.getCol());
    cell.setCellComment(newComment(cellReference.getRow(), cellReference.getCol(), comment));
    ensureLegacyDrawingReference((XSSFSheet) sheet);
  }

  public void clearComment(String address) {
    CellAccess.requireAddress(address);
    CellAccess.optionalCell(sheet, address).ifPresent(CommentAdapter::clearCellComment);
  }

  public static void clearCellComment(Cell cell) {
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

  public static Optional<ResolvedComment> comment(Cell cell) {
    return cell == null ? Optional.empty() : comment(cell.getCellComment());
  }

  public static Optional<ResolvedComment> comment(Comment comment) {
    return comment == null || comment.getString() == null
        ? Optional.empty()
        : comment(comment.getString().getString(), comment.getAuthor(), comment.isVisible());
  }

  public static Optional<ResolvedComment> comment(String text, String author, boolean visible) {
    return text == null || text.isBlank() || author == null || author.isBlank()
        ? Optional.empty()
        : Optional.of(new ResolvedComment(text, author, visible));
  }

  public static Optional<CommentReport> commentSnapshot(Cell cell) {
    return cell == null
        ? Optional.empty()
        : optionalCommentSnapshot(
            Optional.of(cell.getSheet().getWorkbook()), cell.getCellComment());
  }

  public static Optional<CommentReport> commentSnapshot(Comment comment) {
    return optionalCommentSnapshot(Optional.empty(), comment);
  }

  public static void removeCommentFromTable(XSSFSheet sheet, CellAddress address) {
    for (POIXMLDocumentPart relation : sheet.getRelations()) {
      if (relation instanceof Comments comments) {
        comments.removeComment(address);
        return;
      }
    }
  }

  public static void removeCommentShapeIfPresent(XSSFSheet sheet, CellAddress address) {
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

  public static void repairBrokenLegacyDrawingReference(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().isSetLegacyDrawing() && legacyDrawingRelationId(sheet).isEmpty()) {
      sheet.getCTWorksheet().unsetLegacyDrawing();
    }
  }

  public static void ensureLegacyDrawingReference(XSSFSheet sheet) {
    String relationId = vmlDrawingRelationId(sheet).orElse(null);
    if (relationId == null) {
      return;
    }
    if (!sheet.getCTWorksheet().isSetLegacyDrawing()) {
      sheet.getCTWorksheet().addNewLegacyDrawing();
    }
    sheet.getCTWorksheet().getLegacyDrawing().setId(relationId);
  }

  public static Optional<String> legacyDrawingRelationId(XSSFSheet sheet) {
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

  public static Optional<String> vmlDrawingRelationId(XSSFSheet sheet) {
    for (RelationPart relationPart : sheet.getRelationParts()) {
      if (relationPart.getDocumentPart() instanceof XSSFVMLDrawing) {
        return Optional.of(relationPart.getRelationship().getId());
      }
    }
    return Optional.empty();
  }

  private Comment newComment(int rowIndex, int columnIndex, ResolvedComment comment) {
    CellTextLimits.requireSupportedLength(comment.text(), "comment.text");
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
            : StyleAdapter.toPoiRichText(
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

  public List<CellCommentReport> comments() {
    return allUsedComments();
  }

  public List<CellCommentReport> comments(
      List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    return selectedComments(List.copyOf(addresses));
  }

  private List<CellCommentReport> selectedComments(
      List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    List<CellCommentReport> comments = new ArrayList<>();
    for (dev.erst.gridgrind.core.vocabulary.CellAddress address : addresses) {
      Cell cell = CellAccess.cellOrNull(sheet, address.value()).orElse(null);
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
            : StyleAdapter.snapshot(
                (XSSFWorkbook) workbook.orElseThrow(),
                xssfComment.getString(),
                StyleAdapter.snapshotFont(
                    (org.apache.poi.xssf.usermodel.XSSFFont) workbook.get().getFontAt(0)));
    return Optional.of(
        new CommentReport(
            plainComment.orElseThrow().text(),
            plainComment.orElseThrow().author(),
            plainComment.orElseThrow().visible(),
            runs,
            anchor));
  }

  private static final Comparator<CellAddress> CELL_ADDRESS_ORDER =
      Comparator.comparingInt(CellAddress::getRow).thenComparingInt(CellAddress::getColumn);

  public boolean hasPersistedComments() {
    return !sheet.getCellComments().isEmpty()
        || !rawCommentAddresses(commentsTable(sheet).orElse(null)).isEmpty();
  }

  public List<CommentRewriteSnapshot> expectedCommentsAfterInsertColumns(
      int columnIndex, int columnCount) {
    return rewriteCommentsAfterInsertColumns(
        snapshotCommentsForRewrite(), columnIndex, columnCount);
  }

  public List<CommentRewriteSnapshot> expectedCommentsAfterDeleteColumns(ExcelColumnSpan columns) {
    return rewriteCommentsAfterDeleteColumns(snapshotCommentsForRewrite(), columns);
  }

  public List<CommentRewriteSnapshot> expectedCommentsAfterShiftColumns(
      ExcelColumnSpan columns, int delta) {
    return rewriteCommentsAfterShiftColumns(snapshotCommentsForRewrite(), columns, delta);
  }

  public void replaceComments(List<CommentRewriteSnapshot> expectedComments) {
    Objects.requireNonNull(expectedComments, "expectedComments must not be null");
    clearAllPersistedComments();
    for (CommentRewriteSnapshot expectedComment : expectedComments) {
      Objects.requireNonNull(expectedComment, "expectedComments must not contain nulls");
      replaceOneComment(expectedComment);
    }
  }

  public static List<CommentRewriteSnapshot> commentRewriteSnapshots(
      List<CellCommentReport> comments) {
    Objects.requireNonNull(comments, "comments must not be null");
    List<CommentRewriteSnapshot> snapshots = new ArrayList<>(comments.size());
    for (CellCommentReport comment : comments) {
      Objects.requireNonNull(comment, "comments must not contain nulls");
      snapshots.add(CommentRewriteSnapshot.from(comment));
    }
    return List.copyOf(snapshots);
  }

  public static List<CellCommentReport> commentsAfterInsertColumns(
      List<CellCommentReport> comments, int columnIndex, int columnCount) {
    return toReadComments(
        rewriteCommentsAfterInsertColumns(
            commentRewriteSnapshots(comments), columnIndex, columnCount));
  }

  public static List<CellCommentReport> commentsAfterDeleteColumns(
      List<CellCommentReport> comments, ExcelColumnSpan columns) {
    return toReadComments(
        rewriteCommentsAfterDeleteColumns(commentRewriteSnapshots(comments), columns));
  }

  public static List<CellCommentReport> commentsAfterShiftColumns(
      List<CellCommentReport> comments, ExcelColumnSpan columns, int delta) {
    return toReadComments(
        rewriteCommentsAfterShiftColumns(commentRewriteSnapshots(comments), columns, delta));
  }

  private void replaceOneComment(CommentRewriteSnapshot expectedComment) {
    if (expectedComment.isAuthoringCompatible()) {
      setComment(expectedComment.address(), expectedComment.toAuthoringComment());
      return;
    }
    setRawComment(expectedComment);
  }

  private void setRawComment(CommentRewriteSnapshot expectedComment) {
    CellReference reference = new CellReference(expectedComment.address());
    repairBrokenLegacyDrawingReference(sheet);
    Cell cell = getOrCreateCell(sheet, reference.getRow(), reference.getCol());
    cell.setCellComment(newRawComment(reference.getRow(), reference.getCol(), expectedComment));
    ensureLegacyDrawingReference(sheet);
  }

  private XSSFComment newRawComment(
      int rowIndex, int columnIndex, CommentRewriteSnapshot expectedComment) {
    ClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
    Optional<CommentAnchorReport> authoredAnchor = expectedComment.anchor();
    anchor.setRow1(authoredAnchor.map(CommentAnchorReport::firstRow).orElse(rowIndex));
    anchor.setRow2(authoredAnchor.map(CommentAnchorReport::lastRow).orElse(rowIndex + 3));
    anchor.setCol1(authoredAnchor.map(CommentAnchorReport::firstColumn).orElse(columnIndex));
    anchor.setCol2(authoredAnchor.map(CommentAnchorReport::lastColumn).orElse(columnIndex + 3));
    XSSFComment poiComment = (XSSFComment) sheet.createDrawingPatriarch().createCellComment(anchor);
    poiComment.setAuthor(expectedComment.author());
    poiComment.setVisible(expectedComment.visible());
    poiComment.setString(
        expectedComment.runs().isEmpty()
            ? new XSSFRichTextString(expectedComment.text())
            : StyleAdapter.toPoiRichText(
                (XSSFWorkbook) sheet.getWorkbook(),
                toAuthoringRichText(expectedComment.runs().orElseThrow())));
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
    if (commentsTable == null || rawAddresses.isEmpty()) {
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

  private static List<CellCommentReport> toReadComments(List<CommentRewriteSnapshot> comments) {
    List<CellCommentReport> readComments = new ArrayList<>(comments.size());
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

  public static List<CellAddress> rawCommentAddresses(@Nullable CommentsTable commentsTable) {
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

  private static ResolvedRichText toAuthoringRichText(List<RichTextRunReport> runs) {
    return new ResolvedRichText(runs.stream().map(CommentAdapter::toAuthoringRun).toList());
  }

  private static ResolvedRichTextRun toAuthoringRun(RichTextRunReport run) {
    return new ResolvedRichTextRun(run.text(), Optional.of(toAuthoringFont(run.font())));
  }

  private static CellFontPatch toAuthoringFont(CellFontReport font) {
    return new CellFontPatch(
        Optional.of(font.bold()),
        Optional.of(font.italic()),
        Optional.of(font.fontName()),
        Optional.of(font.fontHeight()).map(height -> new FontHeightSpec.Twips(height.twips())),
        font.fontColor().map(ColorAdapter::copyOf),
        Optional.of(font.underline()),
        Optional.of(font.strikeout()));
  }

  public record CommentRewriteSnapshot(
      String address,
      String text,
      String author,
      boolean visible,
      Optional<List<RichTextRunReport>> runs,
      Optional<CommentAnchorReport> anchor) {
    public CommentRewriteSnapshot {
      Objects.requireNonNull(address, "address must not be null");
      text = text == null ? "" : text;
      author = author == null ? "" : author;
      Objects.requireNonNull(runs, "runs must not be null");
      Objects.requireNonNull(anchor, "anchor must not be null");
    }

    public static CommentRewriteSnapshot from(CellCommentReport comment) {
      return new CommentRewriteSnapshot(
          comment.address().value(),
          comment.comment().text(),
          comment.comment().author(),
          comment.comment().visible(),
          comment.comment().runs(),
          comment.comment().anchor());
    }

    public static CommentRewriteSnapshot from(
        CellAddress address, XSSFComment comment, XSSFWorkbook workbook) {
      Optional<CommentAnchorReport> anchor = Optional.empty();
      if (comment.getClientAnchor() instanceof XSSFClientAnchor clientAnchor) {
        anchor =
            Optional.of(
                new CommentAnchorReport(
                    clientAnchor.getCol1(),
                    clientAnchor.getRow1(),
                    clientAnchor.getCol2(),
                    clientAnchor.getRow2()));
      }
      XSSFRichTextString richText = comment.getString();
      Optional<List<RichTextRunReport>> runs =
          richText == null
              ? Optional.empty()
              : StyleAdapter.snapshot(
                  workbook, richText, StyleAdapter.snapshotFont(workbook.getFontAt(0)));
      return new CommentRewriteSnapshot(
          addressString(address.getRow(), address.getColumn()),
          richText == null ? "" : richText.getString(),
          comment.getAuthor(),
          comment.isVisible(),
          runs,
          anchor);
    }

    public CommentRewriteSnapshot at(int rowIndex, int columnIndex) {
      return new CommentRewriteSnapshot(
          addressString(rowIndex, columnIndex), text, author, visible, runs, anchor);
    }

    public boolean isAuthoringCompatible() {
      return !text.isBlank() && !author.isBlank();
    }

    public ResolvedComment toAuthoringComment() {
      return new ResolvedComment(
          text,
          author,
          visible,
          runs.map(CommentAdapter::toAuthoringRichText),
          anchor.map(
              value ->
                  new CommentAnchorSpec(
                      value.firstColumn(), value.firstRow(), value.lastColumn(), value.lastRow())));
    }

    public CellCommentReport toReadComment() {
      return new CellCommentReport(
          new dev.erst.gridgrind.core.vocabulary.CellAddress(address),
          new CommentReport(text, author, visible, runs, anchor));
    }
  }
}
