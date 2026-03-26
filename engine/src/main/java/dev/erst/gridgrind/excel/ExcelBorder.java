package dev.erst.gridgrind.excel;

/**
 * Border patch applied through {@link ExcelCellStyle}, with optional defaults and side overrides.
 */
public record ExcelBorder(
    ExcelBorderSide all,
    ExcelBorderSide top,
    ExcelBorderSide right,
    ExcelBorderSide bottom,
    ExcelBorderSide left) {
  public ExcelBorder {
    if (all == null && top == null && right == null && bottom == null && left == null) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
