package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Authored gradient fill applied through mutable workbook cell styling. */
public sealed interface ExcelGradientFill permits ExcelGradientFill.Linear, ExcelGradientFill.Path {
  /** Ordered authored gradient stops. */
  List<ExcelGradientStop> stops();

  /** Creates one linear authored gradient fill. */
  static Linear linear(Double degree, List<ExcelGradientStop> stops) {
    return new Linear(degree, stops);
  }

  /** Creates one path authored gradient fill. */
  static Path path(
      Double left, Double right, Double top, Double bottom, List<ExcelGradientStop> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Linear gradient authored with one optional degree plus ordered stops. */
  record Linear(Double degree, List<ExcelGradientStop> stops) implements ExcelGradientFill {
    public Linear {
      requireFiniteOrNull(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Path gradient authored with optional edge offsets plus ordered stops. */
  record Path(Double left, Double right, Double top, Double bottom, List<ExcelGradientStop> stops)
      implements ExcelGradientFill {
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
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
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
