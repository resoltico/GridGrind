package dev.erst.gridgrind.contract.dto;

/** One factual cell-relative drawing marker returned by drawing reads. */
public record DrawingMarkerReport(int columnIndex, int rowIndex, int dx, int dy) {
  public DrawingMarkerReport {
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
}
