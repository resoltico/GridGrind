package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Gradient fill payload for cell-style authoring. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellGradientFillInput.Linear.class, name = "LINEAR"),
  @JsonSubTypes.Type(value = CellGradientFillInput.Path.class, name = "PATH")
})
public sealed interface CellGradientFillInput
    permits CellGradientFillInput.Linear, CellGradientFillInput.Path {
  /** Ordered gradient stops; at least two entries are required. */
  List<CellGradientStopInput> stops();

  /** Creates one linear gradient fill. */
  static Linear linear(Double degree, List<CellGradientStopInput> stops) {
    return new Linear(degree, stops);
  }

  /** Creates one path gradient fill. */
  static Path path(
      Double left, Double right, Double top, Double bottom, List<CellGradientStopInput> stops) {
    return new Path(left, right, top, bottom, stops);
  }

  /** Linear gradient defined by one optional degree plus ordered stops. */
  record Linear(Double degree, List<CellGradientStopInput> stops) implements CellGradientFillInput {
    public Linear {
      requireFiniteOrNull(degree, "degree");
      stops = copyStops(stops);
    }
  }

  /** Path gradient defined by optional edge offsets plus ordered stops. */
  record Path(
      Double left, Double right, Double top, Double bottom, List<CellGradientStopInput> stops)
      implements CellGradientFillInput {
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
