package dev.erst.gridgrind.excel;

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

  void setHyperlink(String address, ExcelHyperlink hyperlink) {
    hyperlinkSupport.setHyperlink(address, hyperlink);
  }

  void clearHyperlink(String address) {
    hyperlinkSupport.clearHyperlink(address);
  }

  void setComment(String address, ExcelComment comment) {
    commentSupport.setComment(address, comment);
  }

  void clearComment(String address) {
    commentSupport.clearComment(address);
  }

  ExcelCellMetadataSnapshot metadata(Cell cell) {
    return commentSupport.metadata(cell);
  }

  List<WorkbookSheetResult.CellHyperlink> hyperlinks(ExcelCellSelection selection) {
    return hyperlinkSupport.hyperlinks(selection);
  }

  List<WorkbookSheetResult.CellComment> comments(ExcelCellSelection selection) {
    return commentSupport.comments(selection);
  }

  static void clearCellComment(Cell cell) {
    ExcelSheetCommentSupport.clearCellComment(cell);
  }

  static Optional<ExcelHyperlink> hyperlink(Cell cell) {
    return ExcelSheetHyperlinkSupport.hyperlink(cell);
  }

  static Optional<ExcelHyperlink> hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return ExcelSheetHyperlinkSupport.hyperlink(hyperlink);
  }

  static Optional<ExcelHyperlink> hyperlink(HyperlinkType hyperlinkType, String target) {
    return ExcelSheetHyperlinkSupport.hyperlink(hyperlinkType, target);
  }

  static HyperlinkType toPoi(ExcelHyperlinkType hyperlinkType) {
    return ExcelSheetHyperlinkSupport.toPoi(hyperlinkType);
  }

  static String toPoiTarget(ExcelHyperlink hyperlink) {
    return ExcelSheetHyperlinkSupport.toPoiTarget(hyperlink);
  }

  static Optional<ExcelComment> comment(Cell cell) {
    return ExcelSheetCommentSupport.comment(cell);
  }

  static Optional<ExcelComment> comment(Comment comment) {
    return ExcelSheetCommentSupport.comment(comment);
  }

  static Optional<ExcelComment> comment(String text, String author, boolean visible) {
    return ExcelSheetCommentSupport.comment(text, author, visible);
  }

  static Optional<ExcelCommentSnapshot> commentSnapshot(Cell cell) {
    return ExcelSheetCommentSupport.commentSnapshot(cell);
  }

  static Optional<ExcelCommentSnapshot> commentSnapshot(Comment comment) {
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
}
