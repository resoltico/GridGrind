package dev.erst.gridgrind.protocol.dto;

import java.util.List;
import java.util.Objects;

/** Immutable factual gradient-fill metadata loaded from one cell style. */
public record CellGradientFillReport(
    String type,
    Double degree,
    Double left,
    Double right,
    Double top,
    Double bottom,
    List<CellGradientStopReport> stops) {
  public CellGradientFillReport {
    Objects.requireNonNull(type, "type must not be null");
    if (type.isBlank()) {
      throw new IllegalArgumentException("type must not be blank");
    }
    requireFiniteOrNull(degree, "degree");
    requireFiniteOrNull(left, "left");
    requireFiniteOrNull(right, "right");
    requireFiniteOrNull(top, "top");
    requireFiniteOrNull(bottom, "bottom");
    stops = List.copyOf(Objects.requireNonNull(stops, "stops must not be null"));
    if (stops.isEmpty()) {
      throw new IllegalArgumentException("stops must not be empty");
    }
    for (CellGradientStopReport stop : stops) {
      Objects.requireNonNull(stop, "stops must not contain null values");
    }
  }

  private static void requireFiniteOrNull(Double value, String fieldName) {
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite when provided");
    }
  }
}
