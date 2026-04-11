package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One factual gradient stop preserving its raw position and workbook color semantics. */
public record ExcelGradientStopSnapshot(double position, ExcelColorSnapshot color) {
  public ExcelGradientStopSnapshot {
    if (!Double.isFinite(position) || position < 0.0d || position > 1.0d) {
      throw new IllegalArgumentException("position must be finite and between 0.0 and 1.0");
    }
    Objects.requireNonNull(color, "color must not be null");
  }
}
