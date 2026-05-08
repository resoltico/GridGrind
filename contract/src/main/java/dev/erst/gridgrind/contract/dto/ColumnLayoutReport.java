package dev.erst.gridgrind.contract.dto;

/** Width metadata for one sheet column. */
public record ColumnLayoutReport(
    int columnIndex, double widthCharacters, boolean hidden, int outlineLevel, boolean collapsed) {
  public ColumnLayoutReport {
    if (columnIndex < 0) {
      throw new IllegalArgumentException("columnIndex must not be negative");
    }
    if (!Double.isFinite(widthCharacters) || widthCharacters <= 0.0d) {
      throw new IllegalArgumentException("widthCharacters must be finite and greater than 0");
    }
    if (outlineLevel < 0) {
      throw new IllegalArgumentException("outlineLevel must not be negative");
    }
  }
}
