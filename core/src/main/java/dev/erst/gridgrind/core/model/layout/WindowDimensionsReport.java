package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Published dimensions for one rectangular sheet window. */
@CatalogType(
    id = "WindowDimensionsReport",
    summary = "Published row and column dimensions for one rectangular cell window.",
    displayOrder = 89)
public record WindowDimensionsReport(int rowCount, int columnCount) {
  public WindowDimensionsReport {
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }
  }
}
