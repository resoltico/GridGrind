package dev.erst.gridgrind.jazzer.support;

import java.util.List;
import java.util.function.Supplier;

/** Shared constructor bundle for cartesian chart plot families. */
record ChartCartesianPlotFactories<TAxis, TSeries, TPlot>(
    Supplier<List<TAxis>> categoryAxesFactory,
    Supplier<List<TAxis>> scatterAxesFactory,
    ChartSeriesFactory<TSeries> seriesFactory,
    ChartGroupingPlotFactories.GroupingPlotFactory<TAxis, TSeries, TPlot> areaFactory,
    ChartGroupingPlotFactories.GroupedDepthPlotFactory<TAxis, TSeries, TPlot> area3dFactory,
    ChartBarPlotFactories.BarPlotFactory<TAxis, TSeries, TPlot> barFactory,
    ChartBarPlotFactories.Bar3dPlotFactory<TAxis, TSeries, TPlot> bar3dFactory,
    ChartGroupingPlotFactories.GroupingPlotFactory<TAxis, TSeries, TPlot> lineFactory,
    ChartGroupingPlotFactories.GroupedDepthPlotFactory<TAxis, TSeries, TPlot> line3dFactory,
    ChartPolarPlotFactories.RadarPlotFactory<TAxis, TSeries, TPlot> radarFactory,
    ChartPolarPlotFactories.ScatterPlotFactory<TAxis, TSeries, TPlot> scatterFactory) {}
