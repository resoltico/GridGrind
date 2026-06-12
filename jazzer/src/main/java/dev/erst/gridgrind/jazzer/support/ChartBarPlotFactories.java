package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import java.util.List;
import java.util.Optional;

/** Shared constructor signatures for bar-family chart plots. */
interface ChartBarPlotFactories {

  /** Builds one bar-family plot from sampled bar settings, axes, and series values. */
  @FunctionalInterface
  interface BarPlotFactory<TAxis, TSeries, TPlot> {
    /** Materializes one bar-family plot. */
    TPlot create(
        boolean varyColors,
        ExcelChartBarDirection direction,
        ExcelChartBarGrouping grouping,
        Optional<Integer> gapWidth,
        Optional<Integer> overlap,
        List<TAxis> axes,
        List<TSeries> series);
  }

  /** Builds one 3-D bar-family plot from sampled bar settings, axes, and series values. */
  @FunctionalInterface
  interface Bar3dPlotFactory<TAxis, TSeries, TPlot> {
    /** Materializes one 3-D bar-family plot. */
    TPlot create(
        boolean varyColors,
        ExcelChartBarDirection direction,
        ExcelChartBarGrouping grouping,
        Optional<Integer> gapWidth,
        Optional<Integer> gapDepth,
        Optional<ExcelChartBarShape> shape,
        List<TAxis> axes,
        List<TSeries> series);
  }
}
