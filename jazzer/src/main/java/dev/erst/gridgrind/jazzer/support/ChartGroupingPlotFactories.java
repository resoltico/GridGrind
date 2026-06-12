package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import java.util.List;
import java.util.Optional;

/** Shared constructor signatures for grouped cartesian chart families. */
interface ChartGroupingPlotFactories {

  /** Builds one grouped plot from sampled grouping, axes, and series values. */
  @FunctionalInterface
  interface GroupingPlotFactory<TAxis, TSeries, TPlot> {
    /** Materializes one grouped plot. */
    TPlot create(
        boolean varyColors, ExcelChartGrouping grouping, List<TAxis> axes, List<TSeries> series);
  }

  /** Builds one grouped depth-aware plot from sampled axes and series values. */
  @FunctionalInterface
  interface GroupedDepthPlotFactory<TAxis, TSeries, TPlot> {
    /** Materializes one grouped depth-aware plot. */
    TPlot create(
        boolean varyColors,
        ExcelChartGrouping grouping,
        Optional<Integer> gapDepth,
        List<TAxis> axes,
        List<TSeries> series);
  }
}
