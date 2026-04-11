package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the fill currently applied to a cell. */
public record ExcelCellFillSnapshot(
    ExcelFillPattern pattern,
    ExcelColorSnapshot foregroundColor,
    ExcelColorSnapshot backgroundColor,
    ExcelGradientFillSnapshot gradient) {
  /** Creates a non-gradient fill snapshot. */
  public ExcelCellFillSnapshot(
      ExcelFillPattern pattern,
      ExcelColorSnapshot foregroundColor,
      ExcelColorSnapshot backgroundColor) {
    this(pattern, foregroundColor, backgroundColor, null);
  }

  public ExcelCellFillSnapshot {
    Objects.requireNonNull(pattern, "pattern must not be null");
    if (gradient != null) {
      if (foregroundColor != null || backgroundColor != null) {
        throw new IllegalArgumentException(
            "gradient fills must not also expose foregroundColor or backgroundColor");
      }
    } else {
      if (pattern == ExcelFillPattern.NONE
          && (foregroundColor != null || backgroundColor != null)) {
        throw new IllegalArgumentException("fill pattern NONE does not accept colors");
      }
      if (pattern == ExcelFillPattern.SOLID && backgroundColor != null) {
        throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
      }
    }
  }
}
