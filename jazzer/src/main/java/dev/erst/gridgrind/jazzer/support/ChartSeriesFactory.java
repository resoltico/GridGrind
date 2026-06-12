package dev.erst.gridgrind.jazzer.support;

import java.util.List;

/** Supplies one chart-series family for the current sampled plot. */
@FunctionalInterface
interface ChartSeriesFactory<TSeries> {
  /** Returns the sampled chart series for either pie-like or axis-backed plots. */
  List<TSeries> create(GridGrindFuzzData data, boolean pieChart);
}
