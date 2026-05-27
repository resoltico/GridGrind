package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Layout, merge, pane, zoom, and print operations for one sheet. */
public final class ExcelSheetLayout {
  private final ExcelSheet sheet;
  private final ExcelSheetLayoutSupport layoutSupport;

  ExcelSheetLayout(ExcelSheet sheet, ExcelSheetLayoutSupport layoutSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.layoutSupport = Objects.requireNonNull(layoutSupport, "layoutSupport must not be null");
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  public ExcelSheetLayout mergeCells(String range) {
    layoutSupport.mergeCells(range);
    return this;
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  public ExcelSheetLayout unmergeCells(String range) {
    layoutSupport.unmergeCells(range);
    return this;
  }

  /** Applies one explicit pane state to this sheet. */
  public ExcelSheetLayout setPane(ExcelSheetPane pane) {
    layoutSupport.setPane(pane);
    return this;
  }

  /** Applies one explicit zoom percentage to this sheet. */
  public ExcelSheetLayout setZoom(int zoomPercent) {
    layoutSupport.setZoom(zoomPercent);
    return this;
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  public ExcelSheetLayout setPresentation(ExcelSheetPresentation presentation) {
    layoutSupport.setPresentation(presentation);
    return this;
  }

  /** Applies the provided print layout as the authoritative supported print state. */
  public ExcelSheetLayout setPrintLayout(ExcelPrintLayout printLayout) {
    layoutSupport.setPrintLayout(printLayout);
    return this;
  }

  /** Clears the supported print layout state from this sheet. */
  public ExcelSheetLayout clearPrintLayout() {
    layoutSupport.clearPrintLayout();
    return this;
  }

  /** Returns every merged region currently defined on the sheet. */
  public List<WorkbookSheetResult.MergedRegion> mergedRegions() {
    return layoutSupport.mergedRegions();
  }

  /** Returns layout metadata such as panes, zoom, and visible sizing. */
  public WorkbookSheetResult.SheetLayout snapshot() {
    return layoutSupport.layout(sheet.name());
  }

  /** Returns supported print-layout metadata for this sheet. */
  public ExcelPrintLayout printLayout() {
    return layoutSupport.printLayout();
  }

  /** Returns the full factual print-layout snapshot currently stored for this sheet. */
  public ExcelPrintLayoutSnapshot printLayoutSnapshot() {
    return layoutSupport.printLayoutSnapshot();
  }
}
