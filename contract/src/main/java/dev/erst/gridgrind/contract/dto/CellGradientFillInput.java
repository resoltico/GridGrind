package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Gradient fill payload for cell-style authoring. */
public record CellGradientFillInput(
    String type,
    Double degree,
    Double left,
    Double right,
    Double top,
    Double bottom,
    List<CellGradientStopInput> stops) {
  public CellGradientFillInput {
    type = ExcelGradientFillGeometry.normalizeType(type);
    degree = finiteOrNull(degree, "degree").orElse(null);
    left = finiteOrNull(left, "left").orElse(null);
    right = finiteOrNull(right, "right").orElse(null);
    top = finiteOrNull(top, "top").orElse(null);
    bottom = finiteOrNull(bottom, "bottom").orElse(null);
    ExcelGradientFillGeometry.requireCompatibleGeometry(type, degree, left, right, top, bottom);
    stops = List.copyOf(Objects.requireNonNull(stops, "stops must not be null"));
    if (stops.size() < 2) {
      throw new IllegalArgumentException("stops must contain at least two entries");
    }
    for (CellGradientStopInput stop : stops) {
      Objects.requireNonNull(stop, "stops must not contain null values");
    }
  }

  private static Optional<Double> finiteOrNull(Double value, String fieldName) {
    if (value == null) {
      return Optional.empty();
    }
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    return Optional.of(value);
  }
}
