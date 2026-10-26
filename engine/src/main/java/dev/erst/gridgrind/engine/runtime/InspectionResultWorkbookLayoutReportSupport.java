package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.ArrayList;
import java.util.List;

/** Converts sheet/window/layout/readback snapshots into protocol report records. */
final class InspectionResultWorkbookLayoutReportSupport {
  private InspectionResultWorkbookLayoutReportSupport() {}

  static dev.erst.gridgrind.core.model.layout.WindowReport toWindowReport(
      dev.erst.gridgrind.excel.WorkbookSheetResult.Window window,
      CellReadProjection projection,
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
                          row.rowIndex(), row.cells()))
              .toList());
    }
    List<dev.erst.gridgrind.core.model.value.CellReport> populatedCells = new ArrayList<>();
    for (dev.erst.gridgrind.excel.WorkbookSheetResult.WindowRow row : window.rows()) {
      for (dev.erst.gridgrind.core.model.value.CellReport cell : row.cells()) {
        if (cell instanceof dev.erst.gridgrind.core.model.value.CellReport.BlankReport) {
          continue;
        }
        populatedCells.add(cell);
      }
    }
    return new dev.erst.gridgrind.core.model.layout.WindowReport.Sparse(
        new SheetLabel(window.sheetName()),
        new CellAddress(window.topLeftAddress()),
        dimensions,
        List.copyOf(populatedCells));
  }
}
