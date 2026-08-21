package dev.erst.gridgrind.contract.dto;

import java.util.List;

/** One row inside a rectangular window of cell snapshots. */
public record WindowRowReport(int rowIndex, List<CellReport> cells) {
  public WindowRowReport {
    if (rowIndex < 0) {
      throw new IllegalArgumentException("rowIndex must not be negative");
    }
    cells = WorkbookResultSupport.copyValues(cells, "cells");
  }
}
