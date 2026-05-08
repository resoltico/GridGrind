package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Layout metadata such as pane state, zoom, and visible row and column sizing for one sheet. */
public record SheetLayoutReport(
    String sheetName,
    PaneReport pane,
    int zoomPercent,
    SheetPresentationReport presentation,
    List<ColumnLayoutReport> columns,
    List<RowLayoutReport> rows) {
  public SheetLayoutReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(pane, "pane must not be null");
    Objects.requireNonNull(presentation, "presentation must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireZoomPercent(
        zoomPercent, "zoomPercent");
    columns = GridGrindResponseSupport.copyValues(columns, "columns");
    rows = GridGrindResponseSupport.copyValues(rows, "rows");
  }
}
