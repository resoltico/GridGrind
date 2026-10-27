package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import java.util.List;
import java.util.Objects;

/** Reads reusable introspection facts from one sheet wrapper. */
final class SheetIntrospector {
  /** Returns exact cell snapshots for the provided ordered addresses on one sheet. */
  List<CellReport> cells(XssfSheet sheet, List<String> addresses, CellReadProjection projection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(addresses, "addresses must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    return sheet.cells().snapshotCells(addresses, projection);
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  dev.erst.gridgrind.core.model.layout.WindowReport.Dense window(
      XssfSheet sheet,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      CellReadProjection projection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    return sheet.cells().window(topLeftAddress, rowCount, columnCount, projection);
  }

  /** Returns every merged region currently defined on the sheet. */
  List<MergedRegionReport> mergedRegions(XssfSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout().mergedRegions();
  }

  /** Returns layout metadata such as pane state, zoom, and visible sizing. */
  SheetLayoutReport layout(XssfSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout().snapshot();
  }

  /** Returns supported print-layout metadata for one sheet. */
  PrintLayoutReport printLayout(XssfSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.layout().printLayoutSnapshot();
  }
}
