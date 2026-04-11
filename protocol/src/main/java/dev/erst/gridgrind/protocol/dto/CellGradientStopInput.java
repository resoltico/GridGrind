package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** One gradient stop used by gradient cell-fill authoring. */
public record CellGradientStopInput(double position, ColorInput color) {
  public CellGradientStopInput {
    if (!Double.isFinite(position) || position < 0.0d || position > 1.0d) {
      throw new IllegalArgumentException("position must be finite and between 0.0 and 1.0");
    }
    Objects.requireNonNull(color, "color must not be null");
  }
}
