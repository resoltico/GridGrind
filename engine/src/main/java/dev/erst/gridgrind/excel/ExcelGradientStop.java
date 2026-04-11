package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One authored gradient stop in a mutable workbook fill. */
public record ExcelGradientStop(double position, ExcelColor color) {
  public ExcelGradientStop {
    if (!Double.isFinite(position) || position < 0.0d || position > 1.0d) {
      throw new IllegalArgumentException("position must be finite and between 0.0 and 1.0");
    }
    Objects.requireNonNull(color, "color must not be null");
  }
}
