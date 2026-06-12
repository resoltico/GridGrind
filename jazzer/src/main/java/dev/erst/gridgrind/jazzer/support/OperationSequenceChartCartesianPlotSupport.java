package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartAxisInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import java.util.Optional;

/** Samples cartesian, bar, line, radar, and scatter plot families for paired chart surfaces. */
final class OperationSequenceChartCartesianPlotSupport {
  private OperationSequenceChartCartesianPlotSupport() {}

  static ChartPlotInput nextChartPlotInput(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return nextChartPlot(plotKind, varyColors, data, protocolPlotFactories());
  }

  static ExcelChartDefinition.Plot nextExcelChartPlotDefinition(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return nextChartPlot(plotKind, varyColors, data, excelPlotFactories());
  }

  private static <TAxis, TSeries, TPlot> TPlot nextChartPlot(
      int plotKind,
      boolean varyColors,
      GridGrindFuzzData data,
      ChartCartesianPlotFactories<TAxis, TSeries, TPlot> factories) {
    return switch (plotKind) {
      case 0 ->
          factories
              .areaFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartGrouping(data),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 1 ->
          factories
              .area3dFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartGrouping(data),
                  OperationSequenceChartFactory.nextOptionalInt(data, 0, 500),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 2 ->
          factories
              .barFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartBarDirection(data),
                  OperationSequenceChartFactory.nextChartBarGrouping(data),
                  OperationSequenceChartFactory.nextOptionalInt(data, 0, 500),
                  OperationSequenceChartFactory.nextOptionalInt(data, -100, 100),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 3 ->
          factories
              .bar3dFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartBarDirection(data),
                  OperationSequenceChartFactory.nextChartBarGrouping(data),
                  OperationSequenceChartFactory.nextOptionalInt(data, 0, 500),
                  OperationSequenceChartFactory.nextOptionalInt(data, 0, 500),
                  Optional.of(OperationSequenceChartFactory.nextChartBarShape(data)),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 5 ->
          factories
              .lineFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartGrouping(data),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 6 ->
          factories
              .line3dFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartGrouping(data),
                  OperationSequenceChartFactory.nextOptionalInt(data, 0, 500),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 9 ->
          factories
              .radarFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartRadarStyle(data),
                  factories.categoryAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      case 10 ->
          factories
              .scatterFactory()
              .create(
                  varyColors,
                  OperationSequenceChartFactory.nextChartScatterStyle(data),
                  factories.scatterAxesFactory().get(),
                  factories.seriesFactory().create(data, false));
      default -> throw unsupportedPlotKind(plotKind);
    };
  }

  private static ChartCartesianPlotFactories<ChartAxisInput, ChartSeriesInput, ChartPlotInput>
      protocolPlotFactories() {
    return new ChartCartesianPlotFactories<>(
        OperationSequenceChartFactory::nextChartAxesInputCategory,
        OperationSequenceChartFactory::nextChartAxesInputScatter,
        OperationSequenceChartFactory::nextChartSeriesInputs,
        ChartPlotInput.Area::new,
        ChartPlotInput.Area3D::new,
        ChartPlotInput.Bar::new,
        ChartPlotInput.Bar3D::new,
        ChartPlotInput.Line::new,
        ChartPlotInput.Line3D::new,
        ChartPlotInput.Radar::new,
        ChartPlotInput.Scatter::new);
  }

  private static ChartCartesianPlotFactories<
          ExcelChartDefinition.Axis, ExcelChartDefinition.Series, ExcelChartDefinition.Plot>
      excelPlotFactories() {
    return new ChartCartesianPlotFactories<>(
        OperationSequenceChartFactory::nextExcelChartAxesCategory,
        OperationSequenceChartFactory::nextExcelChartAxesScatter,
        OperationSequenceChartFactory::nextExcelChartSeries,
        ExcelChartDefinition.Area::new,
        ExcelChartDefinition.Area3D::new,
        ExcelChartDefinition.Bar::new,
        ExcelChartDefinition.Bar3D::new,
        ExcelChartDefinition.Line::new,
        ExcelChartDefinition.Line3D::new,
        ExcelChartDefinition.Radar::new,
        ExcelChartDefinition.Scatter::new);
  }

  private static IllegalArgumentException unsupportedPlotKind(int plotKind) {
    return new IllegalArgumentException("Unsupported cartesian plot kind: " + plotKind);
  }
}
