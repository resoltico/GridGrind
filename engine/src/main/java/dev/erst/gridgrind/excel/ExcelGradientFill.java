package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Authored gradient fill applied through mutable workbook cell styling. */
public sealed interface ExcelGradientFill permits ExcelGradientFill.Linear, ExcelGradientFill.Path {
  /** Ordered authored gradient stops. */
  List<ExcelGradientStop> stops();

  /** Creates one linear authored gradient fill. */
  static Linear linear(Optional<Double> degree, List<ExcelGradientStop> stops) {
    return new Linear(degree, stops);
  }

  /** Creates one path authored gradient fill. */
  static Path path(
      Optional<Double> left,
      Optional<Double> right,
      Optional<Double> top,
      Optional<Double> bottom,
      List<ExcelGradientStop> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Linear gradient authored with one optional degree plus ordered stops. */
  record Linear(Optional<Double> degree, List<ExcelGradientStop> stops)
      implements ExcelGradientFill {
    public Linear {
      Objects.requireNonNull(degree, "degree must not be null");
      requireFinite(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Path gradient authored with optional edge offsets plus ordered stops. */
  record Path(
      Optional<Double> left,
      Optional<Double> right,
      Optional<Double> top,
      Optional<Double> bottom,
      List<ExcelGradientStop> stops)
      implements ExcelGradientFill {
    public Path {
      Objects.requireNonNull(left, "left must not be null");
      Objects.requireNonNull(right, "right must not be null");
      Objects.requireNonNull(top, "top must not be null");
      Objects.requireNonNull(bottom, "bottom must not be null");
      requireFinite(left, "left");
      requireFinite(right, "right");
      requireFinite(top, "top");
      requireFinite(bottom, "bottom");
      stops = copyStops(stops);
    }
  }

  private static void requireFinite(Optional<Double> value, String fieldName) {
    value.ifPresent(
        actual -> {
          if (!Double.isFinite(actual)) {
            throw new IllegalArgumentException(fieldName + " must be finite");
          }
        });
  }

  private static List<ExcelGradientStop> copyStops(List<ExcelGradientStop> stops) {
    Objects.requireNonNull(stops, "stops must not be null");
    List<ExcelGradientStop> copy = new ArrayList<>(stops.size());
    for (ExcelGradientStop stop : stops) {
      copy.add(Objects.requireNonNull(stop, "stops must not contain null values"));
    }
    if (copy.size() < 2) {
      throw new IllegalArgumentException("stops must contain at least two entries");
    }
    return List.copyOf(copy);
  }
}
