package dev.erst.gridgrind.excel;

/** One cell-relative marker used by a drawing anchor. */
public record ExcelDrawingMarker(int columnIndex, int rowIndex, int dx, int dy) {
  public ExcelDrawingMarker {
    if (columnIndex < 0) {
      throw new IllegalArgumentException("columnIndex must not be negative");
    }
    if (rowIndex < 0) {
      throw new IllegalArgumentException("rowIndex must not be negative");
    }
    if (dx < 0) {
      throw new IllegalArgumentException("dx must not be negative");
    }
    if (dy < 0) {
      throw new IllegalArgumentException("dy must not be negative");
    }
  }

  /** Creates a marker from raw zero-offset cell coordinates. */
  public ExcelDrawingMarker(int columnIndex, int rowIndex) {
    this(columnIndex, rowIndex, 0, 0);
  }
}
