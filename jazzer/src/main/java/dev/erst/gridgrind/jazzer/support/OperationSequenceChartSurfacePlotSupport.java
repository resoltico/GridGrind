package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.excel.ExcelChartDefinition;

/** Samples surface chart plot families for paired protocol and engine surfaces. */
final class OperationSequenceChartSurfacePlotSupport {
  private OperationSequenceChartSurfacePlotSupport() {}

  static ChartPlotInput nextChartPlotInput(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return switch (plotKind) {
      case 11 ->
          new ChartPlotInput.Surface(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextChartAxesInputSurface(),
              OperationSequenceChartFactory.nextChartSeriesInputs(data, false));
      case 12 ->
          new ChartPlotInput.Surface3D(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextChartAxesInputSurface(),
              OperationSequenceChartFactory.nextChartSeriesInputs(data, false));
      default -> throw unsupportedPlotKind(plotKind);
    };
  }

  static ExcelChartDefinition.Plot nextExcelChartPlotDefinition(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return switch (plotKind) {
      case 11 ->
          new ExcelChartDefinition.Surface(
              varyColors,
              data.consumeBoolean(),
              OperationSequenceChartFactory.nextExcelChartAxesSurface(),
              OperationSequenceChartFactory.nextExcelChartSeries(data, false));
      case 12 ->
          new ExcelChartDefinition.Surface3D(
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
