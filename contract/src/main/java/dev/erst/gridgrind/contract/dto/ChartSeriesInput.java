package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import java.util.Objects;
import java.util.Optional;

/** One authored chart series. */
public record ChartSeriesInput(
    ChartTitleInput title,
    ChartDataSourceInput categories,
    ChartDataSourceInput values,
    Optional<Boolean> smooth,
    Optional<ExcelChartMarkerStyle> markerStyle,
    Optional<Short> markerSize,
    Optional<Long> explosion) {
  public ChartSeriesInput {
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
  public static ChartSeriesInput untitled(
      ChartDataSourceInput categories,
      ChartDataSourceInput values,
      Optional<Boolean> smooth,
      Optional<ExcelChartMarkerStyle> markerStyle,
      Optional<Short> markerSize,
      Optional<Long> explosion) {
    return new ChartSeriesInput(
        new ChartTitleInput.None(), categories, values, smooth, markerStyle, markerSize, explosion);
  }
}
