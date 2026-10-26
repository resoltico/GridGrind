package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/** Resolves source-backed chart titles and series into inline chart definitions. */
final class SourceBackedChartSpecResolver {
  private SourceBackedChartSpecResolver() {}

  static ChartSpec resolveChart(ChartSpec chart, ExecutionInputBindings bindings)
      throws IOException {
    ChartTitleSpec resolvedTitle = resolveChartTitle(chart.title(), bindings);
    List<ChartPlotSpec> resolvedPlots = new ArrayList<>(chart.plots().size());
    boolean changed = !sameReference(resolvedTitle, chart.title());
    for (ChartPlotSpec plot : chart.plots()) {
      ChartPlotSpec resolvedPlot = resolveChartPlot(plot, bindings);
      resolvedPlots.add(resolvedPlot);
      changed |= !sameReference(resolvedPlot, plot);
    }
    return changed
        ? new ChartSpec(
            chart.name(),
            chart.anchor(),
            resolvedTitle,
            chart.legend(),
            chart.displayBlanksAs(),
            chart.plotOnlyVisibleCells(),
            List.copyOf(resolvedPlots))
        : chart;
  }

  private static ChartTitleSpec resolveChartTitle(
      ChartTitleSpec title, ExecutionInputBindings bindings) throws IOException {
    return switch (title) {
      case ChartTitleSpec.None none -> none;
      case ChartTitleSpec.Formula formula -> formula;
      case ChartTitleSpec.Text text -> resolveChartTextTitle(text, bindings);
    };
  }

  private static ChartTitleSpec resolveChartTextTitle(
      ChartTitleSpec.Text text, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedSource =
        SourceBackedPlanResolver.resolveTextSource(text.source(), bindings, true, "chart title");
    return sameReference(resolvedSource, text.source())
        ? text
        : new ChartTitleSpec.Text(resolvedSource);
  }

  private static ChartPlotSpec resolveChartPlot(ChartPlotSpec plot, ExecutionInputBindings bindings)
      throws IOException {
    return switch (plot) {
      case ChartPlotSpec.Area area -> resolveAreaPlot(area, bindings);
      case ChartPlotSpec.Area3D area3D -> resolveArea3DPlot(area3D, bindings);
      case ChartPlotSpec.Bar bar -> resolveBarPlot(bar, bindings);
      case ChartPlotSpec.Bar3D bar3D -> resolveBar3DPlot(bar3D, bindings);
      case ChartPlotSpec.Doughnut doughnut -> resolveDoughnutPlot(doughnut, bindings);
      case ChartPlotSpec.Line line -> resolveLinePlot(line, bindings);
      case ChartPlotSpec.Line3D line3D -> resolveLine3DPlot(line3D, bindings);
      case ChartPlotSpec.Pie pie -> resolvePiePlot(pie, bindings);
      case ChartPlotSpec.Pie3D pie3D -> resolvePie3DPlot(pie3D, bindings);
      case ChartPlotSpec.Radar radar -> resolveRadarPlot(radar, bindings);
      case ChartPlotSpec.Scatter scatter -> resolveScatterPlot(scatter, bindings);
      case ChartPlotSpec.Surface surface -> resolveSurfacePlot(surface, bindings);
      case ChartPlotSpec.Surface3D surface3D -> resolveSurface3DPlot(surface3D, bindings);
    };
  }

  private static ChartPlotSpec.Area resolveAreaPlot(
      ChartPlotSpec.Area plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Area(plot.varyColors(), plot.grouping(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Area3D resolveArea3DPlot(
      ChartPlotSpec.Area3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Area3D(
            plot.varyColors(), plot.grouping(), plot.gapDepth(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Bar resolveBarPlot(
      ChartPlotSpec.Bar plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Bar(
            plot.varyColors(),
            plot.barDirection(),
            plot.grouping(),
            plot.gapWidth(),
            plot.overlap(),
            plot.axes(),
            resolvedSeries);
  }

  private static ChartPlotSpec.Bar3D resolveBar3DPlot(
      ChartPlotSpec.Bar3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Bar3D(
            plot.varyColors(),
            plot.barDirection(),
            plot.grouping(),
            plot.gapDepth(),
            plot.gapWidth(),
            plot.shape(),
            plot.axes(),
            resolvedSeries);
  }

  private static ChartPlotSpec.Doughnut resolveDoughnutPlot(
      ChartPlotSpec.Doughnut plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Doughnut(
            plot.varyColors(), plot.firstSliceAngle(), plot.holeSize(), resolvedSeries);
  }

  private static ChartPlotSpec.Line resolveLinePlot(
      ChartPlotSpec.Line plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Line(plot.varyColors(), plot.grouping(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Line3D resolveLine3DPlot(
      ChartPlotSpec.Line3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Line3D(
            plot.varyColors(), plot.grouping(), plot.gapDepth(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Pie resolvePiePlot(
      ChartPlotSpec.Pie plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Pie(plot.varyColors(), plot.firstSliceAngle(), resolvedSeries);
  }

  private static ChartPlotSpec.Pie3D resolvePie3DPlot(
      ChartPlotSpec.Pie3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Pie3D(plot.varyColors(), resolvedSeries);
  }

  private static ChartPlotSpec.Radar resolveRadarPlot(
      ChartPlotSpec.Radar plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Radar(plot.varyColors(), plot.style(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Scatter resolveScatterPlot(
      ChartPlotSpec.Scatter plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Scatter(plot.varyColors(), plot.style(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Surface resolveSurfacePlot(
      ChartPlotSpec.Surface plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Surface(
            plot.varyColors(), plot.wireframe(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotSpec.Surface3D resolveSurface3DPlot(
      ChartPlotSpec.Surface3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotSpec.Surface3D(
            plot.varyColors(), plot.wireframe(), plot.axes(), resolvedSeries);
  }

  private static List<ChartSeriesSpec> resolveChartSeries(
      List<ChartSeriesSpec> series, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesSpec> resolvedSeries = new ArrayList<>(series.size());
    boolean changed = false;
    for (ChartSeriesSpec value : series) {
      ChartTitleSpec resolvedTitle = resolveChartTitle(value.title(), bindings);
      ChartSeriesSpec resolved = resolveChartSeries(value, resolvedTitle);
      resolvedSeries.add(resolved);
      changed |= !sameReference(resolved, value);
    }
    return changed ? List.copyOf(resolvedSeries) : series;
  }

  private static ChartSeriesSpec resolveChartSeries(
      ChartSeriesSpec series, ChartTitleSpec resolvedTitle) {
    return sameReference(resolvedTitle, series.title())
        ? series
        : new ChartSeriesSpec(
            resolvedTitle,
            series.categories(),
            series.values(),
            series.smooth(),
            series.markerStyle(),
            series.markerSize(),
            series.explosion());
  }
}
