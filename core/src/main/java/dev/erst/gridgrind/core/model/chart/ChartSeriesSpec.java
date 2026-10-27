package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import java.util.Objects;
import java.util.Optional;

/** One authored chart series. */
@CatalogType(
    id = "ChartSeriesSpec",
    summary =
        "One authored chart series with a title plus category and value data sources. smooth, marker, and explosion fields are optional by chart family.",
    displayOrder = 35)
public record ChartSeriesSpec(
    ChartTitleSpec title,
    ChartDataSourceSpec categories,
    ChartDataSourceSpec values,
    Optional<Boolean> smooth,
    Optional<ExcelChartMarkerStyle> markerStyle,
    Optional<Short> markerSize,
    Optional<Long> explosion) {
  public ChartSeriesSpec {
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(categories, "categories must not be null");
    Objects.requireNonNull(values, "values must not be null");
    Objects.requireNonNull(smooth, "smooth must not be null");
    Objects.requireNonNull(markerStyle, "markerStyle must not be null");
    Objects.requireNonNull(markerSize, "markerSize must not be null");
    Objects.requireNonNull(explosion, "explosion must not be null");
    if (markerSize.isPresent() && (markerSize.orElseThrow() < 2 || markerSize.orElseThrow() > 72)) {
      throw new IllegalArgumentException("markerSize must be between 2 and 72");
    }
    if (explosion.isPresent() && explosion.orElseThrow() < 0L) {
      throw new IllegalArgumentException("explosion must not be negative");
    }
  }

  /** Creates one untitled series explicitly. */
  public static ChartSeriesSpec untitled(
      ChartDataSourceSpec categories,
      ChartDataSourceSpec values,
      Optional<Boolean> smooth,
      Optional<ExcelChartMarkerStyle> markerStyle,
      Optional<Short> markerSize,
      Optional<Long> explosion) {
    return new ChartSeriesSpec(
        new ChartTitleSpec.None(), categories, values, smooth, markerStyle, markerSize, explosion);
  }
}
