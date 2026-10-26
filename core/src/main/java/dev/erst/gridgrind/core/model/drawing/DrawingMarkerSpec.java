package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** One cell-relative drawing marker used by authored two-cell anchors. */
@CatalogType(
    id = "DrawingMarkerSpec",
    summary = "One zero-based drawing marker with explicit column, row, and in-cell offsets.",
    displayOrder = 29)
public record DrawingMarkerSpec(int columnIndex, int rowIndex, int dx, int dy) {
  public DrawingMarkerSpec {
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
  public DrawingMarkerSpec(int columnIndex, int rowIndex) {
    this(columnIndex, rowIndex, 0, 0);
  }
}
