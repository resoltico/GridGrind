package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import java.util.Objects;

/** Authored axis definition used by one plot. */
public record ChartAxisInput(
    ExcelChartAxisKind kind,
    ExcelChartAxisPosition position,
    ExcelChartAxisCrosses crosses,
    boolean visible) {
  /** Reads one axis definition with explicit visibility. */
  @JsonCreator
  public ChartAxisInput(
      @JsonProperty("kind") ExcelChartAxisKind kind,
      @JsonProperty("position") ExcelChartAxisPosition position,
      @JsonProperty("crosses") ExcelChartAxisCrosses crosses,
      @JsonProperty("visible") Boolean visible) {
    this(
        kind,
        position,
        crosses,
        Objects.requireNonNull(visible, "visible must not be null").booleanValue());
  }

  public ChartAxisInput {
    Objects.requireNonNull(kind, "kind must not be null");
    Objects.requireNonNull(position, "position must not be null");
    Objects.requireNonNull(crosses, "crosses must not be null");
  }

  /** Creates one visible axis explicitly. */
  public static ChartAxisInput visible(
      ExcelChartAxisKind kind, ExcelChartAxisPosition position, ExcelChartAxisCrosses crosses) {
    return new ChartAxisInput(kind, position, crosses, true);
  }
}
