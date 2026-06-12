package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.excel.ExcelChartDefinition;

/** Routes chart-plot sampling to the bounded plot-family helpers. */
final class OperationSequenceChartPlotSupport {
  private OperationSequenceChartPlotSupport() {}

  static ChartPlotInput nextChartPlotInput(GridGrindFuzzData data) {
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

  static ExcelChartDefinition.Plot nextExcelChartPlotDefinition(GridGrindFuzzData data) {
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
