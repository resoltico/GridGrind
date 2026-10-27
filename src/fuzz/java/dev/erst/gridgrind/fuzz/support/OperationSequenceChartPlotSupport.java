package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;

/** Routes chart-plot sampling to the bounded plot-family helpers. */
final class OperationSequenceChartPlotSupport {
  private OperationSequenceChartPlotSupport() {}

  static ChartPlotSpec nextChartPlotInput(GridGrindFuzzData data) {
    boolean varyColors = data.consumeBoolean();
    int plotKind = nextPlotKind(data);
    return switch (plotKind) {
      case 4, 7, 8 ->
          OperationSequenceChartSlicePlotSupport.nextChartPlotInput(plotKind, varyColors, data);
      case 11, 12 ->
          OperationSequenceChartSurfacePlotSupport.nextChartPlotInput(plotKind, varyColors, data);
      default ->
          OperationSequenceChartCartesianPlotSupport.nextChartPlotInput(plotKind, varyColors, data);
    };
  }

  static ResolvedChart.Plot nextExcelChartPlotDefinition(GridGrindFuzzData data) {
    boolean varyColors = data.consumeBoolean();
    int plotKind = nextPlotKind(data);
    return switch (plotKind) {
      case 4, 7, 8 ->
          OperationSequenceChartSlicePlotSupport.nextExcelChartPlotDefinition(
              plotKind, varyColors, data);
      case 11, 12 ->
          OperationSequenceChartSurfacePlotSupport.nextExcelChartPlotDefinition(
              plotKind, varyColors, data);
      default ->
          OperationSequenceChartCartesianPlotSupport.nextExcelChartPlotDefinition(
              plotKind, varyColors, data);
    };
  }

  private static int nextPlotKind(GridGrindFuzzData data) {
    return OperationSequenceChartFactory.selectorSlot(
            OperationSequenceChartFactory.nextSelectorByte(data))
        % 13;
  }
}
