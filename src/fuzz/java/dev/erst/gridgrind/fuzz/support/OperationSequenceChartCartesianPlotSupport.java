package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import java.util.Optional;

/** Samples cartesian, bar, line, radar, and scatter plot families for paired chart surfaces. */
final class OperationSequenceChartCartesianPlotSupport {
  private OperationSequenceChartCartesianPlotSupport() {}

  static ChartPlotSpec nextChartPlotInput(
      int plotKind, boolean varyColors, GridGrindFuzzData data) {
    return nextChartPlot(plotKind, varyColors, data, protocolPlotFactories());
  }

  static ResolvedChart.Plot nextExcelChartPlotDefinition(
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

  private static ChartCartesianPlotFactories<ChartAxisSpec, ChartSeriesSpec, ChartPlotSpec>
      protocolPlotFactories() {
    return new ChartCartesianPlotFactories<>(
        OperationSequenceChartFactory::nextChartAxesInputCategory,
        OperationSequenceChartFactory::nextChartAxesInputScatter,
        OperationSequenceChartFactory::nextChartSeriesInputs,
        ChartPlotSpec.Area::new,
        ChartPlotSpec.Area3D::new,
        ChartPlotSpec.Bar::new,
        ChartPlotSpec.Bar3D::new,
        ChartPlotSpec.Line::new,
        ChartPlotSpec.Line3D::new,
        ChartPlotSpec.Radar::new,
        ChartPlotSpec.Scatter::new);
  }

  private static ChartCartesianPlotFactories<
          ChartAxisSpec, ResolvedChart.Series, ResolvedChart.Plot>
      excelPlotFactories() {
    return new ChartCartesianPlotFactories<>(
        OperationSequenceChartFactory::nextExcelChartAxesCategory,
        OperationSequenceChartFactory::nextExcelChartAxesScatter,
        OperationSequenceChartFactory::nextExcelChartSeries,
        ResolvedChart.Area::new,
        ResolvedChart.Area3D::new,
        ResolvedChart.Bar::new,
        ResolvedChart.Bar3D::new,
        ResolvedChart.Line::new,
        ResolvedChart.Line3D::new,
        ResolvedChart.Radar::new,
        ResolvedChart.Scatter::new);
  }

  private static IllegalArgumentException unsupportedPlotKind(int plotKind) {
    return new IllegalArgumentException("Unsupported cartesian plot kind: " + plotKind);
  }
}
