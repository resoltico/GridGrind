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
import org.jspecify.annotations.Nullable;

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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(area3D.axes());
    XDDFArea3DChartData areaData =
        (XDDFArea3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.AREA_3D),
                axes.categoryAxis(),
                axes.valueAxis());
    areaData.setVaryColors(area3D.varyColors());
    areaData.setGrouping(ExcelChartPoiBridge.toPoiGrouping(area3D.grouping()));
    area3D.gapDepth().ifPresent(areaData::setGapDepth);
    addSeries(sheet, areaData, area3D.series(), formulaRuntime);
    chart.plot(areaData);
  }

  private static void createBarPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Bar bar,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
    bar.gapWidth().ifPresent(barData::setGapWidth);
    bar.overlap().map(Integer::byteValue).ifPresent(barData::setOverlap);
    addSeries(sheet, barData, bar.series(), formulaRuntime);
    chart.plot(barData);
  }

  private static void createBar3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Bar3D bar3D,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
    bar3D.gapDepth().ifPresent(barData::setGapDepth);
    bar3D.gapWidth().ifPresent(barData::setGapWidth);
    bar3D.shape().map(ExcelChartPoiBridge::toPoiBarShape).ifPresent(barData::setShape);
    addSeries(sheet, barData, bar3D.series(), formulaRuntime);
    chart.plot(barData);
  }

  private static void createDoughnutPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartDefinition.Doughnut doughnut,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    XDDFDoughnutChartData doughnutData =
        (XDDFDoughnutChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.DOUGHNUT), null, null);
    doughnutData.setVaryColors(doughnut.varyColors());
    doughnut.firstSliceAngle().ifPresent(doughnutData::setFirstSliceAngle);
    doughnut.holeSize().ifPresent(doughnutData::setHoleSize);
    addSeries(sheet, doughnutData, doughnut.series(), formulaRuntime);
    chart.plot(doughnutData);
  }

  private static void createLinePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartAxisRegistry axisRegistry,
      ExcelChartDefinition.Line line,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartAxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(line3D.axes());
    XDDFLine3DChartData lineData =
        (XDDFLine3DChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.LINE_3D),
                axes.categoryAxis(),
                axes.valueAxis());
    lineData.setVaryColors(line3D.varyColors());
    lineData.setGrouping(ExcelChartPoiBridge.toPoiGrouping(line3D.grouping()));
    line3D.gapDepth().ifPresent(lineData::setGapDepth);
    addSeries(sheet, lineData, line3D.series(), formulaRuntime);
    chart.plot(lineData);
  }

  private static void createPiePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartDefinition.Pie pie,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    XDDFPieChartData pieData =
        (XDDFPieChartData)
            chart.createData(
                ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.PIE), null, null);
    pieData.setVaryColors(pie.varyColors());
    pie.firstSliceAngle().ifPresent(pieData::setFirstSliceAngle);
    addSeries(sheet, pieData, pie.series(), formulaRuntime);
    chart.plot(pieData);
  }

  private static void createPie3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ExcelChartDefinition.Pie3D pie3D,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
      @Nullable ExcelFormulaRuntime formulaRuntime) {
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
        definition.smooth().ifPresent(lineSeries::setSmooth);
        definition
            .markerStyle()
            .ifPresent(
                markerStyle ->
                    lineSeries.setMarkerStyle(ExcelChartPoiBridge.toPoiMarkerStyle(markerStyle)));
        definition.markerSize().ifPresent(lineSeries::setMarkerSize);
      }
      case XDDFLine3DChartData.Series lineSeries -> {
        definition.smooth().ifPresent(lineSeries::setSmooth);
        definition
            .markerStyle()
            .ifPresent(
                markerStyle ->
                    lineSeries.setMarkerStyle(ExcelChartPoiBridge.toPoiMarkerStyle(markerStyle)));
        definition.markerSize().ifPresent(lineSeries::setMarkerSize);
      }
      case XDDFScatterChartData.Series scatterSeries -> {
        definition.smooth().ifPresent(scatterSeries::setSmooth);
        definition
            .markerStyle()
            .ifPresent(
                markerStyle ->
                    scatterSeries.setMarkerStyle(
                        ExcelChartPoiBridge.toPoiMarkerStyle(markerStyle)));
        definition.markerSize().ifPresent(scatterSeries::setMarkerSize);
      }
      case XDDFPieChartData.Series pieSeries ->
          definition.explosion().ifPresent(pieSeries::setExplosion);
      case XDDFPie3DChartData.Series pieSeries ->
          definition.explosion().ifPresent(pieSeries::setExplosion);
      case XDDFDoughnutChartData.Series doughnutSeries ->
          definition.explosion().ifPresent(doughnutSeries::setExplosion);
      default -> {
        // No extra series-level options for this plot family.
      }
    }
  }
}
