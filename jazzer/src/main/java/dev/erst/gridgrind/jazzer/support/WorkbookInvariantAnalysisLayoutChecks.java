package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.CellCommentReport;
import dev.erst.gridgrind.contract.dto.CellHyperlinkReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.SheetLayoutReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.WindowReport;

/** Owns invariant checks for analysis entry payloads tied to workbook layout and sheet state. */
final class WorkbookInvariantAnalysisLayoutChecks {
  private WorkbookInvariantAnalysisLayoutChecks() {}

  static void requireWindowShape(WindowReport window) {
    WorkbookInvariantChecks.require(
        window.sheetName() != null, "window sheetName must not be null");
    WorkbookInvariantChecks.require(
        !window.sheetName().isBlank(), "window sheetName must not be blank");
    WorkbookInvariantChecks.require(
        window.topLeftAddress() != null, "window topLeftAddress must not be null");
    WorkbookInvariantChecks.require(
        !window.topLeftAddress().isBlank(), "window topLeftAddress must not be blank");
    WorkbookInvariantChecks.require(window.rows() != null, "window rows must not be null");
    WorkbookInvariantChecks.require(
        window.rows().size() == window.rowCount(), "window rows size must match rowCount");
    window
        .rows()
        .forEach(
            row -> {
              WorkbookInvariantChecks.require(
                  row.rowIndex() >= 0, "window row index must not be negative");
              WorkbookInvariantChecks.require(
                  row.cells() != null, "window row cells must not be null");
              WorkbookInvariantChecks.require(
                  row.cells().size() == window.columnCount(),
                  "window row cells size must match columnCount");
              row.cells().forEach(WorkbookInvariantCellSurfaceChecks::requireCellReportShape);
            });
  }

  static void requireHyperlinkEntryShape(CellHyperlinkReport hyperlink) {
    WorkbookInvariantChecks.require(
        hyperlink.address() != null, "hyperlink address must not be null");
    WorkbookInvariantChecks.require(
        !hyperlink.address().isBlank(), "hyperlink address must not be blank");
    WorkbookInvariantChecks.require(
        hyperlink.hyperlink() != null, "hyperlink metadata must not be null");
    WorkbookInvariantCellSurfaceChecks.requireHyperlinkShape(hyperlink.hyperlink());
  }

  static void requireCommentEntryShape(CellCommentReport comment) {
    WorkbookInvariantChecks.require(comment.address() != null, "comment address must not be null");
    WorkbookInvariantChecks.require(
        !comment.address().isBlank(), "comment address must not be blank");
    WorkbookInvariantChecks.require(comment.comment() != null, "comment metadata must not be null");
    WorkbookInvariantCellSurfaceChecks.requireCommentReportShape(comment.comment());
  }

  static void requireSheetLayoutShape(SheetLayoutReport layout) {
    WorkbookInvariantChecks.require(
        layout.sheetName() != null, "layout sheetName must not be null");
    WorkbookInvariantChecks.require(
        !layout.sheetName().isBlank(), "layout sheetName must not be blank");
    WorkbookInvariantChecks.require(layout.pane() != null, "pane must not be null");
    WorkbookInvariantChecks.require(
        layout.zoomPercent() >= 10 && layout.zoomPercent() <= 400,
        "zoomPercent must be between 10 and 400 inclusive");
    switch (layout.pane()) {
      case dev.erst.gridgrind.contract.dto.PaneReport.None _ -> {}
      case dev.erst.gridgrind.contract.dto.PaneReport.Frozen frozen -> {
        WorkbookInvariantChecks.require(
            frozen.splitColumn() >= 0, "splitColumn must not be negative");
        WorkbookInvariantChecks.require(frozen.splitRow() >= 0, "splitRow must not be negative");
        WorkbookInvariantChecks.require(
            frozen.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        WorkbookInvariantChecks.require(frozen.topRow() >= 0, "topRow must not be negative");
      }
      case dev.erst.gridgrind.contract.dto.PaneReport.Split split -> {
        WorkbookInvariantChecks.require(
            split.xSplitPosition() >= 0, "xSplitPosition must not be negative");
        WorkbookInvariantChecks.require(
            split.ySplitPosition() >= 0, "ySplitPosition must not be negative");
        WorkbookInvariantChecks.require(
            split.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        WorkbookInvariantChecks.require(split.topRow() >= 0, "topRow must not be negative");
        WorkbookInvariantChecks.require(split.activePane() != null, "activePane must not be null");
      }
    }
    layout
        .columns()
        .forEach(
            column -> {
              WorkbookInvariantChecks.require(
                  column.columnIndex() >= 0, "columnIndex must not be negative");
              WorkbookInvariantChecks.require(
                  Double.isFinite(column.widthCharacters()) && column.widthCharacters() > 0.0d,
                  "column width must be finite and greater than 0");
            });
    layout
        .rows()
        .forEach(
            row -> {
              WorkbookInvariantChecks.require(row.rowIndex() >= 0, "rowIndex must not be negative");
              WorkbookInvariantChecks.require(
                  Double.isFinite(row.heightPoints()) && row.heightPoints() > 0.0d,
                  "row height must be finite and greater than 0");
            });
  }

  static void requirePrintLayoutShape(PrintLayoutReport layout) {
    WorkbookInvariantChecks.require(
        layout.sheetName() != null, "print layout sheetName must not be null");
    WorkbookInvariantChecks.require(
        !layout.sheetName().isBlank(), "print layout sheetName must not be blank");
    WorkbookInvariantChecks.require(layout.printArea() != null, "printArea must not be null");
    WorkbookInvariantChecks.require(layout.orientation() != null, "orientation must not be null");
    WorkbookInvariantChecks.require(layout.scaling() != null, "scaling must not be null");
    WorkbookInvariantChecks.require(
        layout.repeatingRows() != null, "repeatingRows must not be null");
    WorkbookInvariantChecks.require(
        layout.repeatingColumns() != null, "repeatingColumns must not be null");
    WorkbookInvariantChecks.require(layout.header() != null, "header must not be null");
    WorkbookInvariantChecks.require(layout.footer() != null, "footer must not be null");
    WorkbookInvariantCellSurfaceChecks.requirePrintSetupShape(layout.setup());
  }

  static void requireAutofilterEntryShape(AutofilterEntryReport autofilter) {
    WorkbookInvariantChecks.requireNonBlank(autofilter.range(), "autofilter range");
    WorkbookInvariantChecks.require(
        autofilter.filterColumns() != null, "autofilter filterColumns must not be null");
    autofilter
        .filterColumns()
        .forEach(WorkbookInvariantCellSurfaceChecks::requireAutofilterFilterColumnShape);
    if (autofilter.sortState().isPresent()) {
      WorkbookInvariantCellSurfaceChecks.requireAutofilterSortStateShape(
          autofilter.sortState().orElseThrow());
    }
    switch (autofilter) {
      case AutofilterEntryReport.SheetOwned _ -> {}
      case AutofilterEntryReport.TableOwned tableOwned ->
          WorkbookInvariantChecks.requireNonBlank(tableOwned.tableName(), "autofilter table name");
    }
  }

  static void requireTableEntryShape(TableEntryReport table) {
    WorkbookInvariantChecks.requireNonBlank(table.name(), "table name");
    WorkbookInvariantChecks.requireNonBlank(table.sheetName(), "table sheetName");
    WorkbookInvariantChecks.requireNonBlank(table.range(), "table range");
    WorkbookInvariantChecks.require(
        table.headerRowCount() >= 0, "table headerRowCount must not be negative");
    WorkbookInvariantChecks.require(
        table.totalsRowCount() >= 0, "table totalsRowCount must not be negative");
    WorkbookInvariantChecks.require(
        table.columnNames() != null, "table columnNames must not be null");
    WorkbookInvariantChecks.require(table.columns() != null, "table columns must not be null");
    WorkbookInvariantChecks.require(
        table.columnNames().size() == table.columns().size(),
        "table columnNames size must match columns size");
    table
        .columnNames()
        .forEach(
            columnName ->
                WorkbookInvariantChecks.require(
                    columnName != null, "table column name must not be null"));
    for (int index = 0; index < table.columns().size(); index++) {
      WorkbookInvariantCellSurfaceChecks.requireTableColumnShape(table.columns().get(index));
      WorkbookInvariantChecks.require(
          table.columnNames().get(index).equals(table.columns().get(index).name()),
          "table columnNames must align with columns");
    }
    WorkbookInvariantChecks.require(table.style() != null, "table style must not be null");
    WorkbookInvariantAnalysisFormattingChecks.requireTableStyleShape(table.style());
  }

  static void requirePivotTableShape(PivotTableReport pivotTable) {
    WorkbookInvariantWorkbookSurfaceChecks.requirePivotTableShape(pivotTable);
  }
}
