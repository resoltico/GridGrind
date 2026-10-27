package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Facade that composes hyperlink and comment helpers for one sheet. */
final class ExcelSheetAnnotationSupport {
  private final ExcelSheetHyperlinkSupport hyperlinkSupport;
  private final ExcelSheetCommentSupport commentSupport;

  ExcelSheetAnnotationSupport(Sheet sheet, ExcelDrawingController drawingController) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(drawingController, "drawingController must not be null");
    this.hyperlinkSupport = new ExcelSheetHyperlinkSupport(sheet);
    this.commentSupport = new ExcelSheetCommentSupport(sheet, drawingController);
  }

  void setHyperlink(String address, HyperlinkTarget hyperlink) {
    hyperlinkSupport.setHyperlink(address, hyperlink);
  }

  void clearHyperlink(String address) {
    hyperlinkSupport.clearHyperlink(address);
  }

  void setComment(String address, ResolvedComment comment) {
    commentSupport.setComment(address, comment);
  }

  void clearComment(String address) {
    commentSupport.clearComment(address);
  }

  static void clearCellComment(Cell cell) {
    ExcelSheetCommentSupport.clearCellComment(cell);
  }

  static Optional<HyperlinkTarget> hyperlink(Cell cell) {
    return ExcelSheetHyperlinkSupport.hyperlink(cell);
  }

  static Optional<HyperlinkTarget> hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return ExcelSheetHyperlinkSupport.hyperlink(hyperlink);
  }

  static Optional<HyperlinkTarget> hyperlink(HyperlinkType hyperlinkType, String target) {
    return ExcelSheetHyperlinkSupport.hyperlink(hyperlinkType, target);
  }

  static HyperlinkType toPoi(HyperlinkTarget hyperlinkType) {
    return ExcelSheetHyperlinkSupport.toPoi(hyperlinkType);
  }

  static String toPoiTarget(HyperlinkTarget hyperlink) {
    return ExcelSheetHyperlinkSupport.toPoiTarget(hyperlink);
  }

  static Optional<ResolvedComment> comment(Cell cell) {
    return ExcelSheetCommentSupport.comment(cell);
  }

  static Optional<ResolvedComment> comment(Comment comment) {
    return ExcelSheetCommentSupport.comment(comment);
  }

  static Optional<ResolvedComment> comment(String text, String author, boolean visible) {
    return ExcelSheetCommentSupport.comment(text, author, visible);
  }

  static Optional<CommentReport> commentSnapshot(Cell cell) {
    return ExcelSheetCommentSupport.commentSnapshot(cell);
  }

  static Optional<CommentReport> commentSnapshot(Comment comment) {
    return ExcelSheetCommentSupport.commentSnapshot(comment);
  }

  static void removeCommentFromTable(XSSFSheet sheet, CellAddress address) {
    ExcelSheetCommentSupport.removeCommentFromTable(sheet, address);
  }

  static void removeCommentShapeIfPresent(XSSFSheet sheet, CellAddress address) {
    ExcelSheetCommentSupport.removeCommentShapeIfPresent(sheet, address);
  }

  static void repairBrokenLegacyDrawingReference(XSSFSheet sheet) {
    ExcelSheetCommentSupport.repairBrokenLegacyDrawingReference(sheet);
  }

  static void ensureLegacyDrawingReference(XSSFSheet sheet) {
    ExcelSheetCommentSupport.ensureLegacyDrawingReference(sheet);
  }

  static Optional<String> legacyDrawingRelationId(XSSFSheet sheet) {
    return ExcelSheetCommentSupport.legacyDrawingRelationId(sheet);
  }

  static Optional<String> vmlDrawingRelationId(XSSFSheet sheet) {
    return ExcelSheetCommentSupport.vmlDrawingRelationId(sheet);
  }

  /** Returns metadata for every physically present cell on this bound sheet. */
  List<CellHyperlinkReport> hyperlinks() {
    return hyperlinkSupport.hyperlinks();
  }

  /** Returns metadata for the supplied ordered cell addresses. */
  List<CellHyperlinkReport> hyperlinks(
      List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    return hyperlinkSupport.hyperlinks(addresses);
  }

  /** Returns metadata for every physically present cell on this bound sheet. */
  List<CellCommentReport> comments() {
    return commentSupport.comments();
  }

  /** Returns metadata for the supplied ordered cell addresses. */
  List<CellCommentReport> comments(List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    return commentSupport.comments(addresses);
  }
}
