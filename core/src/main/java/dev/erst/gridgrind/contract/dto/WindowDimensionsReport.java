package dev.erst.gridgrind.contract.dto;

/** Published dimensions for one rectangular sheet window. */
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
