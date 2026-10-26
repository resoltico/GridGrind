package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import java.util.Objects;

/** Authored axis definition used by one plot. */
@CatalogType(
    id = "ChartAxisSpec",
    summary = "One authored chart axis used by a chart plot with explicit visibility state.",
    displayOrder = 34)
public record ChartAxisSpec(
    ExcelChartAxisKind kind,
    ExcelChartAxisPosition position,
    ExcelChartAxisCrosses crosses,
    boolean visible) {

  public ChartAxisSpec {
    Objects.requireNonNull(kind, "kind must not be null");
    Objects.requireNonNull(position, "position must not be null");
    Objects.requireNonNull(crosses, "crosses must not be null");
  }

  /** Creates one visible axis explicitly. */
  public static ChartAxisSpec visible(
      ExcelChartAxisKind kind, ExcelChartAxisPosition position, ExcelChartAxisCrosses crosses) {
    return new ChartAxisSpec(kind, position, crosses, true);
  }
}
