package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** One factual gradient stop loaded from workbook style metadata. */
public record CellGradientStopReport(double position, CellColorReport color) {
  public CellGradientStopReport {
    if (!Double.isFinite(position) || position < 0.0d || position > 1.0d) {
      throw new IllegalArgumentException("position must be finite and between 0.0 and 1.0");
    }
    Objects.requireNonNull(color, "color must not be null");
  }
}
