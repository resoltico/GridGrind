package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Immutable factual gradient-fill payload loaded from one cell style. */
public sealed interface ExcelGradientFillSnapshot
    permits ExcelGradientFillSnapshot.Linear, ExcelGradientFillSnapshot.Path {
  /** Ordered factual gradient stops. */
  List<ExcelGradientStopSnapshot> stops();

  /** Returns one factual linear gradient snapshot. */
  static Linear linear(Double degree, List<ExcelGradientStopSnapshot> stops) {
    return new Linear(degree, stops);
  }

  /** Returns one factual path gradient snapshot. */
  static Path path(
      Double left, Double right, Double top, Double bottom, List<ExcelGradientStopSnapshot> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Factual linear gradient snapshot. */
  record Linear(Double degree, List<ExcelGradientStopSnapshot> stops)
      implements ExcelGradientFillSnapshot {
    public Linear {
      requireFiniteOrNull(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Factual path gradient snapshot. */
  record Path(
      Double left, Double right, Double top, Double bottom, List<ExcelGradientStopSnapshot> stops)
      implements ExcelGradientFillSnapshot {
    public Path {
      requireFiniteOrNull(left, "left");
      requireFiniteOrNull(right, "right");
      requireFiniteOrNull(top, "top");
      requireFiniteOrNull(bottom, "bottom");
      stops = copyStops(stops);
    }
  }

  private static void requireFiniteOrNull(Double value, String fieldName) {
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite when provided");
    }
  }

  private static List<ExcelGradientStopSnapshot> copyStops(List<ExcelGradientStopSnapshot> stops) {
    Objects.requireNonNull(stops, "stops must not be null");
    List<ExcelGradientStopSnapshot> copy = new ArrayList<>(stops.size());
    for (ExcelGradientStopSnapshot stop : stops) {
      copy.add(Objects.requireNonNull(stop, "stops must not contain null values"));
    }
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("stops must not be empty");
    }
    return List.copyOf(copy);
  }
}
