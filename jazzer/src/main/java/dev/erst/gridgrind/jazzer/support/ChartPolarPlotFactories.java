package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import java.util.List;

/** Shared constructor signatures for radar and scatter chart plots. */
interface ChartPolarPlotFactories {

  /** Builds one radar plot from sampled radar settings, axes, and series values. */
  @FunctionalInterface
  interface RadarPlotFactory<TAxis, TSeries, TPlot> {
    /** Materializes one radar plot. */
    TPlot create(
        boolean varyColors,
        ExcelChartRadarStyle radarStyle,
        List<TAxis> axes,
        List<TSeries> series);
  }

  /** Builds one scatter plot from sampled scatter settings, axes, and series values. */
  @FunctionalInterface
  interface ScatterPlotFactory<TAxis, TSeries, TPlot> {
    /** Materializes one scatter plot. */
    TPlot create(
        boolean varyColors,
        ExcelChartScatterStyle scatterStyle,
        List<TAxis> axes,
        List<TSeries> series);
  }
}
