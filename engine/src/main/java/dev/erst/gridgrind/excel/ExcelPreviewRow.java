package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.value.CellReport;
import java.util.List;

/** Lightweight preview of one sheet row for agent-facing inspection. */
public record ExcelPreviewRow(int rowIndex, List<CellReport> cells) {
  public ExcelPreviewRow {
    cells = cells == null ? List.of() : List.copyOf(cells);
  }
}
