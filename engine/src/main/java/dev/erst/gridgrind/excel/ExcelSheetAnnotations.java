package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Cell hyperlink and comment operations for one sheet. */
public final class ExcelSheetAnnotations {
  private final ExcelSheetAnnotationSupport annotationSupport;

  ExcelSheetAnnotations(ExcelSheet sheet, ExcelSheetAnnotationSupport annotationSupport) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    this.annotationSupport =
        Objects.requireNonNull(annotationSupport, "annotationSupport must not be null");
  }

  /** Replaces the hyperlink attached to one cell, creating the cell if necessary. */
  public ExcelSheetAnnotations setHyperlink(String address, ExcelHyperlink hyperlink) {
    annotationSupport.setHyperlink(address, hyperlink);
    return this;
  }

  /** Removes any hyperlink attached to one cell; no-op when the cell does not physically exist. */
  public ExcelSheetAnnotations clearHyperlink(String address) {
    annotationSupport.clearHyperlink(address);
    return this;
  }

  /** Replaces the plain-text comment attached to one cell, creating the cell if necessary. */
  public ExcelSheetAnnotations setComment(String address, ExcelComment comment) {
    annotationSupport.setComment(address, comment);
    return this;
  }

  /** Removes any comment attached to one cell; no-op when the cell does not physically exist. */
  public ExcelSheetAnnotations clearComment(String address) {
    annotationSupport.clearComment(address);
    return this;
  }

  /** Returns hyperlink metadata for the selected cells on this sheet. */
  public List<WorkbookSheetResult.CellHyperlink> hyperlinks(ExcelCellSelection selection) {
    return annotationSupport.hyperlinks(selection);
  }

  /** Returns comment metadata for the selected cells on this sheet. */
  public List<WorkbookSheetResult.CellComment> comments(ExcelCellSelection selection) {
    return annotationSupport.comments(selection);
  }
}
