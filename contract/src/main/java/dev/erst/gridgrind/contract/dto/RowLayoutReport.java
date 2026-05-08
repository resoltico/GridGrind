package dev.erst.gridgrind.contract.dto;

/** Height metadata for one sheet row. */
public record RowLayoutReport(
    int rowIndex, double heightPoints, boolean hidden, int outlineLevel, boolean collapsed) {
  public RowLayoutReport {
    if (rowIndex < 0) {
      throw new IllegalArgumentException("rowIndex must not be negative");
    }
    if (!Double.isFinite(heightPoints) || heightPoints <= 0.0d) {
      throw new IllegalArgumentException("heightPoints must be finite and greater than 0");
    }
    if (outlineLevel < 0) {
      throw new IllegalArgumentException("outlineLevel must not be negative");
    }
  }
}
