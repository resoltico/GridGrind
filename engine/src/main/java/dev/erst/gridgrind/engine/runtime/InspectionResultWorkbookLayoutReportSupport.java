package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import java.util.ArrayList;
import java.util.List;

/** Converts sheet/window/layout/readback snapshots into protocol report records. */
final class InspectionResultWorkbookLayoutReportSupport {
  private InspectionResultWorkbookLayoutReportSupport() {}

  static SheetProtectionReport toSheetProtectionReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection protection) {
    return switch (protection) {
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Unprotected _ ->
          new SheetProtectionReport.Unprotected();
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Protected protectedState ->
          new SheetProtectionReport.Protected(protectedState.settings());
    };
  }

  static dev.erst.gridgrind.core.model.layout.WindowReport toWindowReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.Window window,
      ExcelCellReadProjection projection,
      boolean includeBlanks,
      boolean date1904) {
    dev.erst.gridgrind.core.model.layout.WindowDimensionsReport dimensions =
        new dev.erst.gridgrind.core.model.layout.WindowDimensionsReport(
            window.rowCount(), window.columnCount());
    if (includeBlanks) {
      return new dev.erst.gridgrind.core.model.layout.WindowReport.Dense(
          new SheetLabel(window.sheetName()),
          new CellAddress(window.topLeftAddress()),
          dimensions,
          window.rows().stream()
              .map(
                  row ->
                      new dev.erst.gridgrind.core.model.layout.WindowRowReport(
                          row.rowIndex(),
                          row.cells().stream()
                              .map(
                                  cell ->
                                      InspectionResultCellReportSupport.toCellReport(
                                          cell, projection, date1904))
                              .toList()))
              .toList());
    }
    List<dev.erst.gridgrind.core.model.value.CellReport> populatedCells = new ArrayList<>();
    for (dev.erst.gridgrind.excel.WorkbookSheetResult.WindowRow row : window.rows()) {
      for (dev.erst.gridgrind.excel.ExcelCellSnapshot cell : row.cells()) {
        if (cell instanceof dev.erst.gridgrind.excel.ExcelCellSnapshot.BlankSnapshot) {
          continue;
        }
        populatedCells.add(
            InspectionResultCellReportSupport.toCellReport(cell, projection, date1904));
      }
    }
    return new dev.erst.gridgrind.core.model.layout.WindowReport.Sparse(
        new SheetLabel(window.sheetName()),
        new CellAddress(window.topLeftAddress()),
        dimensions,
        List.copyOf(populatedCells));
  }

  static CellHyperlinkReport toCellHyperlinkReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.CellHyperlink hyperlink) {
    return new CellHyperlinkReport(new CellAddress(hyperlink.address()), hyperlink.hyperlink());
  }

  static CellCommentReport toCellCommentReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.CellComment comment) {
    return new CellCommentReport(new CellAddress(comment.address()), comment.comment());
  }

  static SheetLayoutReport toSheetLayoutReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetLayout layout) {
    return new SheetLayoutReport(
        new SheetLabel(layout.sheetName()),
        toPaneReport(layout.pane()),
        layout.zoomPercent(),
        layout.presentation(),
        layout.columns().stream()
            .map(
                column ->
                    new ColumnLayoutReport(
                        column.columnIndex(),
                        column.widthCharacters(),
                        column.hidden(),
                        column.outlineLevel(),
                        column.collapsed()))
            .toList(),
        layout.rows().stream()
            .map(
                row ->
                    new RowLayoutReport(
                        row.rowIndex(),
                        row.heightPoints(),
                        row.hidden(),
                        row.outlineLevel(),
                        row.collapsed()))
            .toList());
  }

  static PaneReport toPaneReport(ExcelSheetPane pane) {
    return switch (pane) {
      case ExcelSheetPane.None _ -> new PaneReport.None();
      case ExcelSheetPane.Frozen frozen ->
          new PaneReport.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case ExcelSheetPane.Split split ->
          new PaneReport.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }
}
