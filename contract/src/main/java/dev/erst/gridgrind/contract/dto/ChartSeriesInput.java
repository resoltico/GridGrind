package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import java.util.Objects;

/** One authored chart series. */
public record ChartSeriesInput(
    ChartTitleInput title,
    ChartDataSourceInput categories,
    ChartDataSourceInput values,
    Boolean smooth,
    ExcelChartMarkerStyle markerStyle,
    Short markerSize,
    Long explosion) {
  public ChartSeriesInput {
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(categories, "categories must not be null");
    Objects.requireNonNull(values, "values must not be null");
    if (markerSize != null && (markerSize < 2 || markerSize > 72)) {
      throw new IllegalArgumentException("markerSize must be between 2 and 72");
    }
    if (explosion != null && explosion < 0L) {
      throw new IllegalArgumentException("explosion must not be negative");
    }
  }

  /** Creates one untitled series explicitly. */
  public static ChartSeriesInput untitled(
      ChartDataSourceInput categories,
      ChartDataSourceInput values,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      Long explosion) {
    return new ChartSeriesInput(
        new ChartTitleInput.None(), categories, values, smooth, markerStyle, markerSize, explosion);
  }
}
