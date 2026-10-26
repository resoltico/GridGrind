package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Immutable factual gradient-fill metadata loaded from one cell style. */
@Union(discriminator = "type", catalog = true, displayOrder = 60)
public sealed interface CellGradientFillReport
    permits dev.erst.gridgrind.core.model.style.CellGradientFillReport.Linear,
        dev.erst.gridgrind.core.model.style.CellGradientFillReport.Path {
  /** Ordered factual gradient stops. */
  List<CellGradientStopReport> stops();

  /** Returns one factual linear gradient report. */
  static Linear linear(@Nullable Double degree, List<CellGradientStopReport> stops) {
    return new Linear(degree, stops);
  }

  /** Returns one factual path gradient report. */
  static Path path(
      @Nullable Double left,
      @Nullable Double right,
      @Nullable Double top,
      @Nullable Double bottom,
      List<CellGradientStopReport> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Factual linear gradient report. */
  @Variant(id = "LINEAR", summary = "Read one factual linear gradient fill with ordered stops.")
  record Linear(@Nullable Double degree, List<CellGradientStopReport> stops)
      implements CellGradientFillReport {
    public Linear {
      requireFiniteOrNull(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Factual path gradient report. */
  @Variant(
      id = "PATH",
      summary =
          "Read one factual path gradient fill with ordered stops. Edge offsets are omitted when Excel does not persist them.")
  record Path(
      @Nullable Double left,
      @Nullable Double right,
      @Nullable Double top,
      @Nullable Double bottom,
      List<CellGradientStopReport> stops)
      implements CellGradientFillReport {
    public Path {
      requireFiniteOrNull(left, "left");
      requireFiniteOrNull(right, "right");
      requireFiniteOrNull(top, "top");
      requireFiniteOrNull(bottom, "bottom");
      stops = copyStops(stops);
    }
  }

  private static void requireFiniteOrNull(@Nullable Double value, String fieldName) {
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
