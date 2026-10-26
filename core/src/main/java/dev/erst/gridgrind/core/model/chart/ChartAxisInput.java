package dev.erst.gridgrind.core.model.chart;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import java.util.Objects;

/** Authored axis definition used by one plot. */
@CatalogType(
    id = "ChartAxisInput",
    summary = "One authored chart axis used by a chart plot with explicit visibility state.",
    displayOrder = 34)
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
