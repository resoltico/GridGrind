package dev.erst.gridgrind.excel;

import java.util.List;

/** Lightweight preview of one sheet row for agent-facing inspection. */
public record ExcelPreviewRow(int rowIndex, List<ExcelCellSnapshot> cells) {
  public ExcelPreviewRow {
    cells = cells == null ? List.of() : List.copyOf(cells);
  }
}
