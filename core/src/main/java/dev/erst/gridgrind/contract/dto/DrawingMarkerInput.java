package dev.erst.gridgrind.contract.dto;

/** One cell-relative drawing marker used by authored two-cell anchors. */
public record DrawingMarkerInput(int columnIndex, int rowIndex, int dx, int dy) {
  public DrawingMarkerInput {
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

  /** Creates a zero-offset drawing marker. */
  public DrawingMarkerInput(int columnIndex, int rowIndex) {
    this(columnIndex, rowIndex, 0, 0);
  }
}
