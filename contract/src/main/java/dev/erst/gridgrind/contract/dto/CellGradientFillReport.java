package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Immutable factual gradient-fill metadata loaded from one cell style. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellGradientFillReport.Linear.class, name = "LINEAR"),
  @JsonSubTypes.Type(value = CellGradientFillReport.Path.class, name = "PATH")
})
public sealed interface CellGradientFillReport
    permits CellGradientFillReport.Linear, CellGradientFillReport.Path {
  /** Ordered factual gradient stops. */
  List<CellGradientStopReport> stops();

  /** Returns one factual linear gradient report. */
  static Linear linear(Double degree, List<CellGradientStopReport> stops) {
    return new Linear(degree, stops);
  }

  /** Returns one factual path gradient report. */
  static Path path(
      Double left, Double right, Double top, Double bottom, List<CellGradientStopReport> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Factual linear gradient report. */
  record Linear(Double degree, List<CellGradientStopReport> stops)
      implements CellGradientFillReport {
    public Linear {
      requireFiniteOrNull(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Factual path gradient report. */
  record Path(
      Double left, Double right, Double top, Double bottom, List<CellGradientStopReport> stops)
      implements CellGradientFillReport {
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

  private static List<CellGradientStopReport> copyStops(List<CellGradientStopReport> stops) {
    Objects.requireNonNull(stops, "stops must not be null");
    List<CellGradientStopReport> copy = new ArrayList<>(stops.size());
    for (CellGradientStopReport stop : stops) {
      copy.add(Objects.requireNonNull(stop, "stops must not contain null values"));
    }
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("stops must not be empty");
    }
    return List.copyOf(copy);
  }
}
