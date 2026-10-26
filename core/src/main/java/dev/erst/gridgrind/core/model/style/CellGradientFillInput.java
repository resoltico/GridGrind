package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Gradient fill payload for cell-style authoring. */
@Union(discriminator = "type", catalog = true, displayOrder = 56)
public sealed interface CellGradientFillInput
    permits dev.erst.gridgrind.core.model.style.CellGradientFillInput.Linear,
        dev.erst.gridgrind.core.model.style.CellGradientFillInput.Path {
  /** Ordered gradient stops; at least two entries are required. */
  List<CellGradientStopInput> stops();

  /** Creates one linear gradient fill. */
  static Linear linear(Optional<Double> degree, List<CellGradientStopInput> stops) {
    return new Linear(degree, stops);
  }

  /** Creates one path gradient fill. */
  static Path path(
      Optional<Double> left,
      Optional<Double> right,
      Optional<Double> top,
      Optional<Double> bottom,
      List<CellGradientStopInput> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Linear gradient defined by one optional degree plus ordered stops. */
  @Variant(
      id = "LINEAR",
      summary =
          "Write one linear gradient fill with ordered stops. degree is optional when Excel's default angle is acceptable.")
  record Linear(Optional<Double> degree, List<CellGradientStopInput> stops)
      implements CellGradientFillInput {
    public Linear {
      Objects.requireNonNull(degree, "degree must not be null");
      requireFinite(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Path gradient defined by optional edge offsets plus ordered stops. */
  @Variant(
      id = "PATH",
      summary =
          "Write one path gradient fill with ordered stops. Each edge offset is optional and omitted offsets preserve Excel's default path geometry.")
  record Path(
      Optional<Double> left,
      Optional<Double> right,
      Optional<Double> top,
      Optional<Double> bottom,
      List<CellGradientStopInput> stops)
      implements CellGradientFillInput {
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

  private static List<CellGradientStopInput> copyStops(List<CellGradientStopInput> stops) {
    Objects.requireNonNull(stops, "stops must not be null");
    List<CellGradientStopInput> copy = new ArrayList<>(stops.size());
    for (CellGradientStopInput stop : stops) {
      copy.add(Objects.requireNonNull(stop, "stops must not contain null values"));
    }
    if (copy.size() < 2) {
      throw new IllegalArgumentException("stops must contain at least two entries");
    }
    return List.copyOf(copy);
  }
}
