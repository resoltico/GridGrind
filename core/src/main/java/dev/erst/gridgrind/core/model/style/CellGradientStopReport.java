package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;

/** One factual gradient stop loaded from workbook style metadata. */
@CatalogType(
    id = "CellGradientStopReport",
    summary = "Exact gradient stop report.",
    displayOrder = 84)
public record CellGradientStopReport(double position, CellColorReport color) {
  public CellGradientStopReport {
    if (!Double.isFinite(position) || position < 0.0d || position > 1.0d) {
      throw new IllegalArgumentException("position must be finite and between 0.0 and 1.0");
    }
    Objects.requireNonNull(color, "color must not be null");
  }
}
