package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import java.util.List;

/** One row inside a rectangular window of cell snapshots. */
@CatalogType(
    id = "WindowRowReport",
    summary = "One dense window row preserving row index and ordered cell payloads.",
    displayOrder = 90)
public record WindowRowReport(int rowIndex, List<CellReport> cells) {
  public WindowRowReport {
    if (rowIndex < 0) {
      throw new IllegalArgumentException("rowIndex must not be negative");
    }
    cells = WorkbookResultSupport.copyValues(cells, "cells");
  }
}
