package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;

/** One gradient stop used by gradient cell-fill authoring. */
@CatalogType(
    id = "CellGradientStopInput",
    summary = "One gradient stop with a normalized position between 0.0 and 1.0.",
    displayOrder = 48)
public record CellGradientStopInput(double position, ColorInput color) {
  public CellGradientStopInput {
    if (!Double.isFinite(position) || position < 0.0d || position > 1.0d) {
      throw new IllegalArgumentException("position must be finite and between 0.0 and 1.0");
    }
    Objects.requireNonNull(color, "color must not be null");
  }
}
