package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Reads reusable introspection facts from one sheet wrapper. */
final class ExcelSheetIntrospector {
  /** Returns exact cell snapshots for the provided ordered addresses on one sheet. */
  List<ExcelCellSnapshot> cells(ExcelSheet sheet, List<String> addresses) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(addresses, "addresses must not be null");
    return sheet.snapshotCells(addresses);
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  WorkbookSheetResult.Window window(
      ExcelSheet sheet, String topLeftAddress, int rowCount, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.window(topLeftAddress, rowCount, columnCount);
  }

  /** Returns every merged region currently defined on the sheet. */
  List<WorkbookSheetResult.MergedRegion> mergedRegions(ExcelSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.mergedRegions();
  }

  /** Returns hyperlink metadata for the selected cells on one sheet. */
  List<WorkbookSheetResult.CellHyperlink> hyperlinks(
      ExcelSheet sheet, ExcelCellSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return sheet.hyperlinks(selection);
  }

  /** Returns comment metadata for the selected cells on one sheet. */
  List<WorkbookSheetResult.CellComment> comments(ExcelSheet sheet, ExcelCellSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return sheet.comments(selection);
  }

  /** Returns layout metadata such as pane state, zoom, and visible sizing. */
  WorkbookSheetResult.SheetLayout layout(ExcelSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout();
  }

  /** Returns supported print-layout metadata for one sheet. */
  ExcelPrintLayoutSnapshot printLayout(ExcelSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.printLayoutSnapshot();
  }
}
