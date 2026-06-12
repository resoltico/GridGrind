package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.excel.ExcelChartDefinition;

/** Samples pie-like chart plot families for paired protocol and engine surfaces. */
final class OperationSequenceChartSlicePlotSupport {
  private OperationSequenceChartSlicePlotSupport() {}

  static ChartPlotInput nextChartPlotInput(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return switch (plotKind) {
      case 4 ->
          new ChartPlotInput.Doughnut(
              varyColors,
              OperationSequenceChartFactory.nextOptionalInt(data, 0, 360),
              OperationSequenceChartFactory.nextOptionalInt(data, 10, 90),
              OperationSequenceChartFactory.nextChartSeriesInputs(data, true));
      case 7 ->
          new ChartPlotInput.Pie(
              varyColors,
              OperationSequenceChartFactory.nextOptionalInt(data, 0, 360),
              OperationSequenceChartFactory.nextChartSeriesInputs(data, true));
      case 8 ->
          new ChartPlotInput.Pie3D(
              varyColors, OperationSequenceChartFactory.nextChartSeriesInputs(data, true));
      default -> throw unsupportedPlotKind(plotKind);
    };
  }

  static ExcelChartDefinition.Plot nextExcelChartPlotDefinition(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return switch (plotKind) {
      case 4 ->
          new ExcelChartDefinition.Doughnut(
              varyColors,
              OperationSequenceChartFactory.nextOptionalInt(data, 0, 360),
              OperationSequenceChartFactory.nextOptionalInt(data, 10, 90),
              OperationSequenceChartFactory.nextExcelChartSeries(data, true));
      case 7 ->
          new ExcelChartDefinition.Pie(
              varyColors,
              OperationSequenceChartFactory.nextOptionalInt(data, 0, 360),
              OperationSequenceChartFactory.nextExcelChartSeries(data, true));
      case 8 ->
          new ExcelChartDefinition.Pie3D(
              varyColors, OperationSequenceChartFactory.nextExcelChartSeries(data, true));
      default -> throw unsupportedPlotKind(plotKind);
    };
  }

  private static IllegalArgumentException unsupportedPlotKind(int plotKind) {
    return new IllegalArgumentException("Unsupported slice plot kind: " + plotKind);
  }
}
