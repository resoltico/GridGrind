package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/** Resolves source-backed chart titles and series into inline chart definitions. */
final class SourceBackedChartInputResolver {
  private SourceBackedChartInputResolver() {}

  static ChartInput resolveChart(ChartInput chart, ExecutionInputBindings bindings)
      throws IOException {
    ChartTitleInput resolvedTitle = resolveChartTitle(chart.title(), bindings);
    List<ChartPlotInput> resolvedPlots = new ArrayList<>(chart.plots().size());
    boolean changed = !sameReference(resolvedTitle, chart.title());
    for (ChartPlotInput plot : chart.plots()) {
      ChartPlotInput resolvedPlot = resolveChartPlot(plot, bindings);
      resolvedPlots.add(resolvedPlot);
      changed |= !sameReference(resolvedPlot, plot);
    }
    return changed
        ? new ChartInput(
            chart.name(),
            chart.anchor(),
            resolvedTitle,
            chart.legend(),
            chart.displayBlanksAs(),
            chart.plotOnlyVisibleCells(),
            List.copyOf(resolvedPlots))
        : chart;
  }

  private static ChartTitleInput resolveChartTitle(
      ChartTitleInput title, ExecutionInputBindings bindings) throws IOException {
    return switch (title) {
      case ChartTitleInput.None none -> none;
      case ChartTitleInput.Formula formula -> formula;
      case ChartTitleInput.Text text -> resolveChartTextTitle(text, bindings);
    };
  }

  private static ChartTitleInput resolveChartTextTitle(
      ChartTitleInput.Text text, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedSource =
        SourceBackedPlanResolver.resolveTextSource(text.source(), bindings, true, "chart title");
    return sameReference(resolvedSource, text.source())
        ? text
        : new ChartTitleInput.Text(resolvedSource);
  }

  private static ChartPlotInput resolveChartPlot(
      ChartPlotInput plot, ExecutionInputBindings bindings) throws IOException {
    return switch (plot) {
      case ChartPlotInput.Area area -> resolveAreaPlot(area, bindings);
      case ChartPlotInput.Area3D area3D -> resolveArea3DPlot(area3D, bindings);
      case ChartPlotInput.Bar bar -> resolveBarPlot(bar, bindings);
      case ChartPlotInput.Bar3D bar3D -> resolveBar3DPlot(bar3D, bindings);
      case ChartPlotInput.Doughnut doughnut -> resolveDoughnutPlot(doughnut, bindings);
      case ChartPlotInput.Line line -> resolveLinePlot(line, bindings);
      case ChartPlotInput.Line3D line3D -> resolveLine3DPlot(line3D, bindings);
      case ChartPlotInput.Pie pie -> resolvePiePlot(pie, bindings);
      case ChartPlotInput.Pie3D pie3D -> resolvePie3DPlot(pie3D, bindings);
      case ChartPlotInput.Radar radar -> resolveRadarPlot(radar, bindings);
      case ChartPlotInput.Scatter scatter -> resolveScatterPlot(scatter, bindings);
      case ChartPlotInput.Surface surface -> resolveSurfacePlot(surface, bindings);
      case ChartPlotInput.Surface3D surface3D -> resolveSurface3DPlot(surface3D, bindings);
    };
  }

  private static ChartPlotInput.Area resolveAreaPlot(
      ChartPlotInput.Area plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Area(plot.varyColors(), plot.grouping(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Area3D resolveArea3DPlot(
      ChartPlotInput.Area3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Area3D(
            plot.varyColors(), plot.grouping(), plot.gapDepth(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Bar resolveBarPlot(
      ChartPlotInput.Bar plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Bar(
            plot.varyColors(),
            plot.barDirection(),
            plot.grouping(),
            plot.gapWidth(),
            plot.overlap(),
            plot.axes(),
            resolvedSeries);
  }

  private static ChartPlotInput.Bar3D resolveBar3DPlot(
      ChartPlotInput.Bar3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Bar3D(
            plot.varyColors(),
            plot.barDirection(),
            plot.grouping(),
            plot.gapDepth(),
            plot.gapWidth(),
            plot.shape(),
            plot.axes(),
            resolvedSeries);
  }

  private static ChartPlotInput.Doughnut resolveDoughnutPlot(
      ChartPlotInput.Doughnut plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Doughnut(
            plot.varyColors(), plot.firstSliceAngle(), plot.holeSize(), resolvedSeries);
  }

  private static ChartPlotInput.Line resolveLinePlot(
      ChartPlotInput.Line plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Line(plot.varyColors(), plot.grouping(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Line3D resolveLine3DPlot(
      ChartPlotInput.Line3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Line3D(
            plot.varyColors(), plot.grouping(), plot.gapDepth(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Pie resolvePiePlot(
      ChartPlotInput.Pie plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Pie(plot.varyColors(), plot.firstSliceAngle(), resolvedSeries);
  }

  private static ChartPlotInput.Pie3D resolvePie3DPlot(
      ChartPlotInput.Pie3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Pie3D(plot.varyColors(), resolvedSeries);
  }

  private static ChartPlotInput.Radar resolveRadarPlot(
      ChartPlotInput.Radar plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Radar(plot.varyColors(), plot.style(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Scatter resolveScatterPlot(
      ChartPlotInput.Scatter plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Scatter(plot.varyColors(), plot.style(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Surface resolveSurfacePlot(
      ChartPlotInput.Surface plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Surface(
            plot.varyColors(), plot.wireframe(), plot.axes(), resolvedSeries);
  }

  private static ChartPlotInput.Surface3D resolveSurface3DPlot(
      ChartPlotInput.Surface3D plot, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = resolveChartSeries(plot.series(), bindings);
    return sameReference(resolvedSeries, plot.series())
        ? plot
        : new ChartPlotInput.Surface3D(
            plot.varyColors(), plot.wireframe(), plot.axes(), resolvedSeries);
  }

  private static List<ChartSeriesInput> resolveChartSeries(
      List<ChartSeriesInput> series, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = new ArrayList<>(series.size());
    boolean changed = false;
    for (ChartSeriesInput value : series) {
      ChartTitleInput resolvedTitle = resolveChartTitle(value.title(), bindings);
      ChartSeriesInput resolved = resolveChartSeries(value, resolvedTitle);
      resolvedSeries.add(resolved);
      changed |= !sameReference(resolved, value);
    }
    return changed ? List.copyOf(resolvedSeries) : series;
  }

  private static ChartSeriesInput resolveChartSeries(
      ChartSeriesInput series, ChartTitleInput resolvedTitle) {
    return sameReference(resolvedTitle, series.title())
        ? series
        : new ChartSeriesInput(
            resolvedTitle,
            series.categories(),
            series.values(),
            series.smooth(),
            series.markerStyle(),
            series.markerSize(),
            series.explosion());
  }
}
