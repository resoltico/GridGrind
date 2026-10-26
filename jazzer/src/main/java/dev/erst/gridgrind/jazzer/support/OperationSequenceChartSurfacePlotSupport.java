package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;

/** Samples surface chart plot families for paired protocol and engine surfaces. */
final class OperationSequenceChartSurfacePlotSupport {
  private OperationSequenceChartSurfacePlotSupport() {}

  static ChartPlotSpec nextChartPlotInput(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return switch (plotKind) {
      case 11 ->
          new ChartPlotSpec.Surface(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextChartAxesInputSurface(),
              OperationSequenceChartFactory.nextChartSeriesInputs(data, false));
      case 12 ->
          new ChartPlotSpec.Surface3D(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextChartAxesInputSurface(),
              OperationSequenceChartFactory.nextChartSeriesInputs(data, false));
      default -> throw unsupportedPlotKind(plotKind);
    };
  }

  static ResolvedChart.Plot nextExcelChartPlotDefinition(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return switch (plotKind) {
      case 11 ->
          new ResolvedChart.Surface(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextExcelChartAxesSurface(),
              OperationSequenceChartFactory.nextExcelChartSeries(data, false));
      case 12 ->
          new ResolvedChart.Surface3D(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextExcelChartAxesSurface(),
              OperationSequenceChartFactory.nextExcelChartSeries(data, false));
      default -> throw unsupportedPlotKind(plotKind);
    };
  }

  private static IllegalArgumentException unsupportedPlotKind(int plotKind) {
    return new IllegalArgumentException("Unsupported surface plot kind: " + plotKind);
  }
}
