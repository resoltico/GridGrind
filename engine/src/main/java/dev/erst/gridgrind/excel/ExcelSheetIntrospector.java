package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import java.util.List;
import java.util.Objects;

/** Reads reusable introspection facts from one sheet wrapper. */
final class ExcelSheetIntrospector {
  /** Returns exact cell snapshots for the provided ordered addresses on one sheet. */
  List<CellReport> cells(ExcelSheet sheet, List<String> addresses, CellReadProjection projection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(addresses, "addresses must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    return sheet.cells().snapshotCells(addresses, projection);
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  WorkbookSheetResult.Window window(
      ExcelSheet sheet,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      CellReadProjection projection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    return sheet.cells().window(topLeftAddress, rowCount, columnCount, projection);
  }

  /** Returns every merged region currently defined on the sheet. */
  List<MergedRegionReport> mergedRegions(ExcelSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout().mergedRegions();
  }

  /** Returns hyperlink metadata for the selected cells on one sheet. */
  List<CellHyperlinkReport> hyperlinks(ExcelSheet sheet, ExcelCellSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return sheet.annotations().hyperlinks(selection);
  }

  /** Returns comment metadata for the selected cells on one sheet. */
  List<CellCommentReport> comments(ExcelSheet sheet, ExcelCellSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return sheet.annotations().comments(selection);
  }

  /** Returns layout metadata such as pane state, zoom, and visible sizing. */
  SheetLayoutReport layout(ExcelSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout().snapshot();
  }

  /** Returns supported print-layout metadata for one sheet. */
  PrintLayoutReport printLayout(ExcelSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout().printLayoutSnapshot();
  }
}
