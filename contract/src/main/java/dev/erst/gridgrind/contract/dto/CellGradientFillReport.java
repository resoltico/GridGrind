package dev.erst.gridgrind.contract.dto;

import java.util.ArrayList;
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
    stops = copyValues(stops, "stops");
    if (stops.isEmpty()) {
      throw new IllegalArgumentException("stops must not be empty");
    }
  }

  private static void requireFiniteOrNull(Double value, String fieldName) {
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite when provided");
    }
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copy);
  }
}
