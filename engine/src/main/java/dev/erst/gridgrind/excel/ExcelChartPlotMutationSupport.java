package dev.erst.gridgrind.excel;

import java.util.List;
import org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Writes one authored plot into a POI chart while reusing shared chart helpers. */
final class ExcelChartPlotMutationSupport {
  private ExcelChartPlotMutationSupport() {}

  static void createPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Plot plot) {
    createPlot(sheet, chart, axisRegistry, plot, null);
  }

  static void createPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Plot plot,
      ExcelFormulaRuntime formulaRuntime) {
    switch (plot) {
      case ExcelChartDefinition.Area area ->
          createAreaPlot(sheet, chart, axisRegistry, area, formulaRuntime);
      case ExcelChartDefinition.Area3D area3D ->
          createArea3DPlot(sheet, chart, axisRegistry, area3D, formulaRuntime);
      case ExcelChartDefinition.Bar bar ->
          createBarPlot(sheet, chart, axisRegistry, bar, formulaRuntime);
      case ExcelChartDefinition.Bar3D bar3D ->
          createBar3DPlot(sheet, chart, axisRegistry, bar3D, formulaRuntime);
      case ExcelChartDefinition.Doughnut doughnut ->
          createDoughnutPlot(sheet, chart, doughnut, formulaRuntime);
      case ExcelChartDefinition.Line line ->
          createLinePlot(sheet, chart, axisRegistry, line, formulaRuntime);
      case ExcelChartDefinition.Line3D line3D ->
          createLine3DPlot(sheet, chart, axisRegistry, line3D, formulaRuntime);
      case ExcelChartDefinition.Pie pie -> createPiePlot(sheet, chart, pie, formulaRuntime);
      case ExcelChartDefinition.Pie3D pie3D -> createPie3DPlot(sheet, chart, pie3D, formulaRuntime);
      case ExcelChartDefinition.Radar radar ->
          createRadarPlot(sheet, chart, axisRegistry, radar, formulaRuntime);
      case ExcelChartDefinition.Scatter scatter ->
          createScatterPlot(sheet, chart, axisRegistry, scatter, formulaRuntime);
      case ExcelChartDefinition.Surface surface ->
          createSurfacePlot(sheet, chart, axisRegistry, surface, formulaRuntime);
      case ExcelChartDefinition.Surface3D surface3D ->
          createSurface3DPlot(sheet, chart, axisRegistry, surface3D, formulaRuntime);
    }
  }

  private static void createAreaPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Area area,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(area.axes());
    XDDFAreaChartData areaData =
        (XDDFAreaChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.AREA),
                axes.categoryAxis(),
                axes.valueAxis());
    areaData.setVaryColors(area.varyColors());
    areaData.setGrouping(ExcelChartPoiBridge.toPoiGrouping(area.grouping()));
    addSeries(sheet, areaData, area.series(), formulaRuntime);
    chart.plot(areaData);
  }

  private static void createArea3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Area3D area3D,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(area3D.axes());
    XDDFArea3DChartData areaData =
        (XDDFArea3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.AREA_3D),
                axes.categoryAxis(),
                axes.valueAxis());
    areaData.setVaryColors(area3D.varyColors());
    areaData.setGrouping(ExcelChartPoiBridge.toPoiGrouping(area3D.grouping()));
    areaData.setGapDepth(area3D.gapDepth());
    addSeries(sheet, areaData, area3D.series(), formulaRuntime);
    chart.plot(areaData);
  }

  private static void createBarPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Bar bar,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(bar.axes());
    XDDFBarChartData barData =
        (XDDFBarChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.BAR),
                axes.categoryAxis(),
                axes.valueAxis());
    barData.setVaryColors(bar.varyColors());
    barData.setBarDirection(ExcelChartPoiBridge.toPoiBarDirection(bar.barDirection()));
    barData.setBarGrouping(ExcelChartPoiBridge.toPoiBarGrouping(bar.grouping()));
    barData.setGapWidth(bar.gapWidth());
    barData.setOverlap(bar.overlap() == null ? null : bar.overlap().byteValue());
    addSeries(sheet, barData, bar.series(), formulaRuntime);
    chart.plot(barData);
  }

  private static void createBar3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Bar3D bar3D,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(bar3D.axes());
    XDDFBar3DChartData barData =
        (XDDFBar3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.BAR_3D),
                axes.categoryAxis(),
                axes.valueAxis());
    barData.setVaryColors(bar3D.varyColors());
    barData.setBarDirection(ExcelChartPoiBridge.toPoiBarDirection(bar3D.barDirection()));
    barData.setBarGrouping(ExcelChartPoiBridge.toPoiBarGrouping(bar3D.grouping()));
    barData.setGapDepth(bar3D.gapDepth());
    barData.setGapWidth(bar3D.gapWidth());
    if (bar3D.shape() != null) {
      barData.setShape(ExcelChartPoiBridge.toPoiBarShape(bar3D.shape()));
    }
    addSeries(sheet, barData, bar3D.series(), formulaRuntime);
    chart.plot(barData);
  }

  private static void createDoughnutPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartDefinition.Doughnut doughnut,
      ExcelFormulaRuntime formulaRuntime) {
    XDDFDoughnutChartData doughnutData =
        (XDDFDoughnutChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.DOUGHNUT), null, null);
    doughnutData.setVaryColors(doughnut.varyColors());
    doughnutData.setFirstSliceAngle(doughnut.firstSliceAngle());
    doughnutData.setHoleSize(doughnut.holeSize());
    addSeries(sheet, doughnutData, doughnut.series(), formulaRuntime);
    chart.plot(doughnutData);
  }

  private static void createLinePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Line line,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(line.axes());
    XDDFLineChartData lineData =
        (XDDFLineChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.LINE),
                axes.categoryAxis(),
                axes.valueAxis());
    lineData.setVaryColors(line.varyColors());
    lineData.setGrouping(ExcelChartPoiBridge.toPoiGrouping(line.grouping()));
    addSeries(sheet, lineData, line.series(), formulaRuntime);
    chart.plot(lineData);
  }

  private static void createLine3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Line3D line3D,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(line3D.axes());
    XDDFLine3DChartData lineData =
        (XDDFLine3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.LINE_3D),
                axes.categoryAxis(),
                axes.valueAxis());
    lineData.setVaryColors(line3D.varyColors());
    lineData.setGrouping(ExcelChartPoiBridge.toPoiGrouping(line3D.grouping()));
    lineData.setGapDepth(line3D.gapDepth());
    addSeries(sheet, lineData, line3D.series(), formulaRuntime);
    chart.plot(lineData);
  }

  private static void createPiePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartDefinition.Pie pie,
      ExcelFormulaRuntime formulaRuntime) {
    XDDFPieChartData pieData =
        (XDDFPieChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.PIE), null, null);
    pieData.setVaryColors(pie.varyColors());
    pieData.setFirstSliceAngle(pie.firstSliceAngle());
    addSeries(sheet, pieData, pie.series(), formulaRuntime);
    chart.plot(pieData);
  }

  private static void createPie3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartDefinition.Pie3D pie3D,
      ExcelFormulaRuntime formulaRuntime) {
    XDDFPie3DChartData pieData =
        (XDDFPie3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.PIE_3D), null, null);
    pieData.setVaryColors(pie3D.varyColors());
    addSeries(sheet, pieData, pie3D.series(), formulaRuntime);
    chart.plot(pieData);
  }

  private static void createRadarPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Radar radar,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(radar.axes());
    XDDFRadarChartData radarData =
        (XDDFRadarChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.RADAR),
                axes.categoryAxis(),
                axes.valueAxis());
    radarData.setVaryColors(radar.varyColors());
    radarData.setStyle(ExcelChartPoiBridge.toPoiRadarStyle(radar.style()));
    addSeries(sheet, radarData, radar.series(), formulaRuntime);
    chart.plot(radarData);
  }

  private static void createScatterPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Scatter scatter,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.ScatterAxes axes = axisRegistry.scatterAxes(scatter.axes());
    XDDFScatterChartData scatterData =
        (XDDFScatterChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.SCATTER),
                axes.xAxis(),
                axes.yAxis());
    scatterData.setVaryColors(scatter.varyColors());
    scatterData.setStyle(ExcelChartPoiBridge.toPoiScatterStyle(scatter.style()));
    addSeries(sheet, scatterData, scatter.series(), formulaRuntime);
    chart.plot(scatterData);
  }

  private static void createSurfacePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Surface surface,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.SurfaceAxes axes = axisRegistry.surfaceAxes(surface.axes());
    XDDFSurfaceChartData surfaceData =
        (XDDFSurfaceChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.SURFACE),
                axes.categoryAxis(),
                axes.valueAxis());
    surfaceData.setVaryColors(surface.varyColors());
    surfaceData.defineSeriesAxis(axes.seriesAxis());
    surfaceData.setWireframe(surface.wireframe());
    addSeries(sheet, surfaceData, surface.series(), formulaRuntime);
    chart.plot(surfaceData);
  }

  private static void createSurface3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Surface3D surface3D,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.SurfaceAxes axes = axisRegistry.surfaceAxes(surface3D.axes());
    XDDFSurface3DChartData surfaceData =
        (XDDFSurface3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.SURFACE_3D),
                axes.categoryAxis(),
                axes.valueAxis());
    surfaceData.setVaryColors(surface3D.varyColors());
    surfaceData.defineSeriesAxis(axes.seriesAxis());
    surfaceData.setWireframe(surface3D.wireframe());
    addSeries(sheet, surfaceData, surface3D.series(), formulaRuntime);
    chart.plot(surfaceData);
  }

  private static void addSeries(
      XSSFSheet sheet,
      XDDFChartData data,
      List<ExcelChartDefinition.Series> definitions,
      ExcelFormulaRuntime formulaRuntime) {
    for (ExcelChartDefinition.Series definition : definitions) {
      XDDFChartData.Series series =
          data.addSeries(
              ExcelChartSourceSupport.toCategoryDataSource(
                  sheet, definition.categories(), formulaRuntime),
              ExcelChartSourceSupport.toValueDataSource(
                  sheet, definition.values(), formulaRuntime));
      applySeriesTitle(sheet, series, definition.title(), formulaRuntime);
      applySeriesOptions(series, definition);
    }
  }

  private static void applySeriesTitle(
      XSSFSheet sheet,
      XDDFChartData.Series series,
      ExcelChartDefinition.Title title,
      ExcelFormulaRuntime formulaRuntime) {
    switch (ExcelChartMutationSupport.prepareSeriesTitle(sheet, title, formulaRuntime)) {
      case PreparedSeriesTitleNone _ -> {
        // Leave the series title unset.
      }
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  static void applySeriesOptions(
      XDDFChartData.Series series, ExcelChartDefinition.Series definition) {
    switch (series) {
      case XDDFLineChartData.Series lineSeries -> {
        if (definition.smooth() != null) {
          lineSeries.setSmooth(definition.smooth());
        }
        if (definition.markerStyle() != null) {
          lineSeries.setMarkerStyle(ExcelChartPoiBridge.toPoiMarkerStyle(definition.markerStyle()));
        }
        if (definition.markerSize() != null) {
          lineSeries.setMarkerSize(definition.markerSize());
        }
      }
      case XDDFLine3DChartData.Series lineSeries -> {
        if (definition.smooth() != null) {
          lineSeries.setSmooth(definition.smooth());
        }
        if (definition.markerStyle() != null) {
          lineSeries.setMarkerStyle(ExcelChartPoiBridge.toPoiMarkerStyle(definition.markerStyle()));
        }
        if (definition.markerSize() != null) {
          lineSeries.setMarkerSize(definition.markerSize());
        }
      }
      case XDDFScatterChartData.Series scatterSeries -> {
        if (definition.smooth() != null) {
          scatterSeries.setSmooth(definition.smooth());
        }
        if (definition.markerStyle() != null) {
          scatterSeries.setMarkerStyle(
              ExcelChartPoiBridge.toPoiMarkerStyle(definition.markerStyle()));
        }
        if (definition.markerSize() != null) {
          scatterSeries.setMarkerSize(definition.markerSize());
        }
      }
      case XDDFPieChartData.Series pieSeries -> {
        if (definition.explosion() != null) {
          pieSeries.setExplosion(definition.explosion());
        }
      }
      case XDDFPie3DChartData.Series pieSeries -> {
        if (definition.explosion() != null) {
          pieSeries.setExplosion(definition.explosion());
        }
      }
      case XDDFDoughnutChartData.Series doughnutSeries -> {
        if (definition.explosion() != null) {
          doughnutSeries.setExplosion(definition.explosion());
        }
      }
      default -> {
        // No extra series-level options for this plot family.
      }
    }
  }
}
