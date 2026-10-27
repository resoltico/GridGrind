package dev.erst.gridgrind.engine.charts;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import javax.xml.namespace.QName;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.DisplayBlanks;
import org.apache.poi.xddf.usermodel.chart.Grouping;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.RadarStyle;
import org.apache.poi.xddf.usermodel.chart.ScatterStyle;
import org.apache.poi.xddf.usermodel.chart.Shape;
import org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDateAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSeriesAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTArea3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBar3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLine3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPie3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScatterChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurface3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurfaceChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Owns chart validation, creation, source binding, titles, plot readback and chart relationships.
 */
// The unexported adapter owns one complete feature; method count does not split its boundary.
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired", "PMD.ExcessivePublicCount"})
public final class ChartAdapter {
  private ChartAdapter() {}

  public static ExcelChartAxisKind axisKind(XDDFChartAxis axis) {
    Objects.requireNonNull(axis, "axis must not be null");
    return switch (axis) {
      case XDDFCategoryAxis _ -> ExcelChartAxisKind.CATEGORY;
      case XDDFDateAxis _ -> ExcelChartAxisKind.DATE;
      case XDDFSeriesAxis _ -> ExcelChartAxisKind.SERIES;
      case XDDFValueAxis _ -> ExcelChartAxisKind.VALUE;
      default ->
          throw new IllegalArgumentException("Unsupported chart axis family: " + axis.getClass());
    };
  }

  private static XDDFChartAxis createAxis(
      XDDFChart chart, ExcelChartAxisKind axisKind, ExcelChartAxisPosition position) {
    Objects.requireNonNull(chart, "chart must not be null");
    Objects.requireNonNull(axisKind, "axisKind must not be null");
    Objects.requireNonNull(position, "position must not be null");
    return switch (axisKind) {
      case CATEGORY -> chart.createCategoryAxis(toPoiAxisPosition(position));
      case DATE -> chart.createDateAxis(toPoiAxisPosition(position));
      case SERIES -> chart.createSeriesAxis(toPoiAxisPosition(position));
      case VALUE -> chart.createValueAxis(toPoiAxisPosition(position));
    };
  }

  public static String plotTypeToken(XDDFChartData chartData) {
    Objects.requireNonNull(chartData, "chartData must not be null");
    return switch (chartData) {
      case XDDFAreaChartData _ -> "AREA";
      case XDDFArea3DChartData _ -> "AREA_3D";
      case XDDFBarChartData _ -> "BAR";
      case XDDFBar3DChartData _ -> "BAR_3D";
      case XDDFDoughnutChartData _ -> "DOUGHNUT";
      case XDDFLineChartData _ -> "LINE";
      case XDDFLine3DChartData _ -> "LINE_3D";
      case XDDFPieChartData _ -> "PIE";
      case XDDFPie3DChartData _ -> "PIE_3D";
      case XDDFRadarChartData _ -> "RADAR";
      case XDDFScatterChartData _ -> "SCATTER";
      case XDDFSurfaceChartData _ -> "SURFACE";
      case XDDFSurface3DChartData _ -> "SURFACE_3D";
      default -> canonicalPlotTypeToken(chartData.getClass().getSimpleName());
    };
  }

  public static String canonicalPlotTypeToken(String simpleName) {
    Objects.requireNonNull(simpleName, "simpleName must not be null");
    if (simpleName.startsWith("XDDF") && simpleName.endsWith("ChartData")) {
      return simpleName
          .substring(4, simpleName.length() - "ChartData".length())
          .toUpperCase(Locale.ROOT);
    }
    return simpleName.toUpperCase(Locale.ROOT);
  }

  private static final Map<String, ExcelChartMarkerStyle> FROM_SYMBOL_TOKEN =
      Map.ofEntries(
          Map.entry("CIRCLE", ExcelChartMarkerStyle.CIRCLE),
          Map.entry("DASH", ExcelChartMarkerStyle.DASH),
          Map.entry("DIAMOND", ExcelChartMarkerStyle.DIAMOND),
          Map.entry("DOT", ExcelChartMarkerStyle.DOT),
          Map.entry("NONE", ExcelChartMarkerStyle.NONE),
          Map.entry("PICTURE", ExcelChartMarkerStyle.PICTURE),
          Map.entry("PLUS", ExcelChartMarkerStyle.PLUS),
          Map.entry("SQUARE", ExcelChartMarkerStyle.SQUARE),
          Map.entry("STAR", ExcelChartMarkerStyle.STAR),
          Map.entry("TRIANGLE", ExcelChartMarkerStyle.TRIANGLE),
          Map.entry("X", ExcelChartMarkerStyle.X));

  public static ExcelChartMarkerStyle fromPoi(MarkerStyle style) {
    return ExcelChartMarkerStyle.valueOf(style.name());
  }

  public static MarkerStyle toPoi(ExcelChartMarkerStyle style) {
    return MarkerStyle.valueOf(style.name());
  }

  public static Optional<ExcelChartMarkerStyle> fromSymbolToken(String symbolToken) {
    if (symbolToken == null || symbolToken.isBlank()) {
      return Optional.empty();
    }
    return Optional.ofNullable(FROM_SYMBOL_TOKEN.get(symbolToken.toUpperCase(Locale.ROOT)));
  }

  public static void validateChart(XSSFSheet sheet, ResolvedChart definition) {
    validateChart(sheet, definition, null);
  }

  public static void validateChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    validateChartTitle(sheet, definition.title(), formulaRuntime);
    for (ResolvedChart.Plot plot : definition.plots()) {
      validatePlot(sheet, plot, formulaRuntime);
    }
  }

  public static void createChart(XSSFSheet sheet, ResolvedChart definition) {
    createChart(sheet, definition, null);
  }

  public static void createChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        DrawingAdapter.toPoiAnchor(drawing, definition.anchor());
    XSSFChart chart = drawing.createChart(anchor);
    chart.getGraphicFrame().setName(definition.name());
    applyCommonChartState(sheet, chart, definition, formulaRuntime);

    AxisRegistry axisRegistry = new AxisRegistry(chart);
    for (ResolvedChart.Plot plot : definition.plots()) {
      createPlot(sheet, chart, axisRegistry, plot, formulaRuntime);
    }
  }

  public static ChartReport.Title prepareSeriesTitle(XSSFSheet sheet, ResolvedChartTitle title) {
    return prepareSeriesTitle(sheet, title, null);
  }

  public static ChartReport.Title prepareSeriesTitle(
      XSSFSheet sheet, ResolvedChartTitle title, @Nullable FormulaRuntime formulaRuntime) {
    return switch (title) {
      case ChartTitleSpec.None _ -> new ChartReport.Title.None();
      case ResolvedChartTitle.Text text -> new ChartReport.Title.Text(text.text());
      case ChartTitleSpec.Formula formula -> {
        CellReference reference =
            resolveSingleCellReference(
                sheet, normalizeFormula(formula.formula()), "Series title formula");
        yield new ChartReport.Title.Formula(
            reference.formatAsString(), scalarText(sheet, reference, formulaRuntime));
      }
    };
  }

  private static void validateChartTitle(
      XSSFSheet sheet, ResolvedChartTitle title, @Nullable FormulaRuntime formulaRuntime) {
    if (title instanceof ChartTitleSpec.Formula formula) {
      CellReference reference =
          resolveSingleCellReference(
              sheet, normalizeFormula(formula.formula()), "Chart title formula");
      scalarText(sheet, reference, formulaRuntime);
    }
  }

  private static void validatePlot(
      XSSFSheet sheet, ResolvedChart.Plot plot, @Nullable FormulaRuntime formulaRuntime) {
    for (ResolvedChart.Series series : plotSeries(plot)) {
      toCategoryDataSource(sheet, series.categories(), formulaRuntime);
      toValueDataSource(sheet, series.values(), formulaRuntime);
      prepareSeriesTitle(sheet, series.title(), formulaRuntime);
    }
  }

  private static List<ResolvedChart.Series> plotSeries(ResolvedChart.Plot plot) {
    return switch (plot) {
      case ResolvedChart.Area area -> area.series();
      case ResolvedChart.Area3D area3D -> area3D.series();
      case ResolvedChart.Bar bar -> bar.series();
      case ResolvedChart.Bar3D bar3D -> bar3D.series();
      case ResolvedChart.Doughnut doughnut -> doughnut.series();
      case ResolvedChart.Line line -> line.series();
      case ResolvedChart.Line3D line3D -> line3D.series();
      case ResolvedChart.Pie pie -> pie.series();
      case ResolvedChart.Pie3D pie3D -> pie3D.series();
      case ResolvedChart.Radar radar -> radar.series();
      case ResolvedChart.Scatter scatter -> scatter.series();
      case ResolvedChart.Surface surface -> surface.series();
      case ResolvedChart.Surface3D surface3D -> surface3D.series();
    };
  }

  private static void applyCommonChartState(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChart definition,
      @Nullable FormulaRuntime formulaRuntime) {
    applyChartTitle(sheet, chart, definition.title(), formulaRuntime);
    applyChartLegend(chart, definition.legend());
    chart.displayBlanksAs(toPoiDisplayBlanks(definition.displayBlanksAs()));
    chart.setPlotOnlyVisibleCells(definition.plotOnlyVisibleCells());
  }

  private static void applyChartTitle(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChartTitle title,
      @Nullable FormulaRuntime formulaRuntime) {
    switch (title) {
      case ChartTitleSpec.None _ -> chart.removeTitle();
      case ResolvedChartTitle.Text text -> chart.setTitleText(text.text());
      case ChartTitleSpec.Formula formula -> {
        CellReference reference =
            resolveSingleCellReference(
                sheet, normalizeFormula(formula.formula()), "Chart title formula");
        applyChartTitleFormula(chart, scalarText(sheet, reference, formulaRuntime), reference);
      }
    }
  }

  public static void applyChartTitleFormula(
      XSSFChart chart, String cachedText, CellReference reference) {
    CTTitle title =
        chart.getCTChart().isSetTitle()
            ? chart.getCTChart().getTitle()
            : chart.getCTChart().addNewTitle();
    CTTx text = title.isSetTx() ? title.getTx() : title.addNewTx();
    if (text.isSetRich()) {
      text.unsetRich();
    }
    CTStrRef strRef = text.isSetStrRef() ? text.getStrRef() : text.addNewStrRef();
    strRef.setF(reference.formatAsString());
    writeStringReferenceCache(strRef, cachedText);
  }

  private static void writeStringReferenceCache(CTStrRef reference, String cachedText) {
    if (reference.isSetStrCache()) {
      reference.unsetStrCache();
    }
    CTStrData cache = reference.addNewStrCache();
    cache.addNewPtCount().setVal(1);
    cache.addNewPt().setIdx(0);
    cache.getPtArray(0).setV(cachedText);
  }

  private static void applyChartLegend(XSSFChart chart, ChartLegendSpec legend) {
    switch (legend) {
      case ChartLegendSpec.Hidden _ -> chart.deleteLegend();
      case ChartLegendSpec.Visible visible ->
          chart.getOrAddLegend().setPosition(toPoiLegendPosition(visible.position()));
    }
  }

  public static void createPlot(
      XSSFSheet sheet, XSSFChart chart, AxisRegistry axisRegistry, ResolvedChart.Plot plot) {
    createPlot(sheet, chart, axisRegistry, plot, null);
  }

  public static void createPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Plot plot,
      @Nullable FormulaRuntime formulaRuntime) {
    switch (plot) {
      case ResolvedChart.Area area ->
          createAreaPlot(sheet, chart, axisRegistry, area, formulaRuntime);
      case ResolvedChart.Area3D area3D ->
          createArea3DPlot(sheet, chart, axisRegistry, area3D, formulaRuntime);
      case ResolvedChart.Bar bar -> createBarPlot(sheet, chart, axisRegistry, bar, formulaRuntime);
      case ResolvedChart.Bar3D bar3D ->
          createBar3DPlot(sheet, chart, axisRegistry, bar3D, formulaRuntime);
      case ResolvedChart.Doughnut doughnut ->
          createDoughnutPlot(sheet, chart, doughnut, formulaRuntime);
      case ResolvedChart.Line line ->
          createLinePlot(sheet, chart, axisRegistry, line, formulaRuntime);
      case ResolvedChart.Line3D line3D ->
          createLine3DPlot(sheet, chart, axisRegistry, line3D, formulaRuntime);
      case ResolvedChart.Pie pie -> createPiePlot(sheet, chart, pie, formulaRuntime);
      case ResolvedChart.Pie3D pie3D -> createPie3DPlot(sheet, chart, pie3D, formulaRuntime);
      case ResolvedChart.Radar radar ->
          createRadarPlot(sheet, chart, axisRegistry, radar, formulaRuntime);
      case ResolvedChart.Scatter scatter ->
          createScatterPlot(sheet, chart, axisRegistry, scatter, formulaRuntime);
      case ResolvedChart.Surface surface ->
          createSurfacePlot(sheet, chart, axisRegistry, surface, formulaRuntime);
      case ResolvedChart.Surface3D surface3D ->
          createSurface3DPlot(sheet, chart, axisRegistry, surface3D, formulaRuntime);
    }
  }

  private static void createAreaPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Area area,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(area.axes());
    XDDFAreaChartData areaData =
        (XDDFAreaChartData)
            chart.createData(ChartTypes.AREA, axes.categoryAxis(), axes.valueAxis());
    areaData.setVaryColors(area.varyColors());
    areaData.setGrouping(toPoiGrouping(area.grouping()));
    addSeries(sheet, areaData, area.series(), formulaRuntime);
    chart.plot(areaData);
  }

  private static void createArea3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Area3D area3D,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(area3D.axes());
    XDDFArea3DChartData areaData =
        (XDDFArea3DChartData)
            chart.createData(ChartTypes.AREA3D, axes.categoryAxis(), axes.valueAxis());
    areaData.setVaryColors(area3D.varyColors());
    areaData.setGrouping(toPoiGrouping(area3D.grouping()));
    area3D.gapDepth().ifPresent(areaData::setGapDepth);
    addSeries(sheet, areaData, area3D.series(), formulaRuntime);
    chart.plot(areaData);
  }

  private static void createBarPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Bar bar,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(bar.axes());
    XDDFBarChartData barData =
        (XDDFBarChartData) chart.createData(ChartTypes.BAR, axes.categoryAxis(), axes.valueAxis());
    barData.setVaryColors(bar.varyColors());
    barData.setBarDirection(toPoiBarDirection(bar.barDirection()));
    barData.setBarGrouping(toPoiBarGrouping(bar.grouping()));
    bar.gapWidth().ifPresent(barData::setGapWidth);
    bar.overlap().map(Integer::byteValue).ifPresent(barData::setOverlap);
    addSeries(sheet, barData, bar.series(), formulaRuntime);
    chart.plot(barData);
  }

  private static void createBar3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Bar3D bar3D,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(bar3D.axes());
    XDDFBar3DChartData barData =
        (XDDFBar3DChartData)
            chart.createData(ChartTypes.BAR3D, axes.categoryAxis(), axes.valueAxis());
    barData.setVaryColors(bar3D.varyColors());
    barData.setBarDirection(toPoiBarDirection(bar3D.barDirection()));
    barData.setBarGrouping(toPoiBarGrouping(bar3D.grouping()));
    bar3D.gapDepth().ifPresent(barData::setGapDepth);
    bar3D.gapWidth().ifPresent(barData::setGapWidth);
    bar3D.shape().map(ChartAdapter::toPoiBarShape).ifPresent(barData::setShape);
    addSeries(sheet, barData, bar3D.series(), formulaRuntime);
    chart.plot(barData);
  }

  private static void createDoughnutPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChart.Doughnut doughnut,
      @Nullable FormulaRuntime formulaRuntime) {
    XDDFDoughnutChartData doughnutData =
        (XDDFDoughnutChartData) chart.createData(ChartTypes.DOUGHNUT, null, null);
    doughnutData.setVaryColors(doughnut.varyColors());
    doughnut.firstSliceAngle().ifPresent(doughnutData::setFirstSliceAngle);
    doughnut.holeSize().ifPresent(doughnutData::setHoleSize);
    addSeries(sheet, doughnutData, doughnut.series(), formulaRuntime);
    chart.plot(doughnutData);
  }

  private static void createLinePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Line line,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(line.axes());
    XDDFLineChartData lineData =
        (XDDFLineChartData)
            chart.createData(ChartTypes.LINE, axes.categoryAxis(), axes.valueAxis());
    lineData.setVaryColors(line.varyColors());
    lineData.setGrouping(toPoiGrouping(line.grouping()));
    addSeries(sheet, lineData, line.series(), formulaRuntime);
    chart.plot(lineData);
  }

  private static void createLine3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Line3D line3D,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(line3D.axes());
    XDDFLine3DChartData lineData =
        (XDDFLine3DChartData)
            chart.createData(ChartTypes.LINE3D, axes.categoryAxis(), axes.valueAxis());
    lineData.setVaryColors(line3D.varyColors());
    lineData.setGrouping(toPoiGrouping(line3D.grouping()));
    line3D.gapDepth().ifPresent(lineData::setGapDepth);
    addSeries(sheet, lineData, line3D.series(), formulaRuntime);
    chart.plot(lineData);
  }

  private static void createPiePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChart.Pie pie,
      @Nullable FormulaRuntime formulaRuntime) {
    XDDFPieChartData pieData = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
    pieData.setVaryColors(pie.varyColors());
    pie.firstSliceAngle().ifPresent(pieData::setFirstSliceAngle);
    addSeries(sheet, pieData, pie.series(), formulaRuntime);
    chart.plot(pieData);
  }

  private static void createPie3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChart.Pie3D pie3D,
      @Nullable FormulaRuntime formulaRuntime) {
    XDDFPie3DChartData pieData =
        (XDDFPie3DChartData) chart.createData(ChartTypes.PIE3D, null, null);
    pieData.setVaryColors(pie3D.varyColors());
    addSeries(sheet, pieData, pie3D.series(), formulaRuntime);
    chart.plot(pieData);
  }

  private static void createRadarPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Radar radar,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.CategoryValueAxes axes = axisRegistry.categoryValueAxes(radar.axes());
    XDDFRadarChartData radarData =
        (XDDFRadarChartData)
            chart.createData(ChartTypes.RADAR, axes.categoryAxis(), axes.valueAxis());
    radarData.setVaryColors(radar.varyColors());
    radarData.setStyle(toPoiRadarStyle(radar.style()));
    addSeries(sheet, radarData, radar.series(), formulaRuntime);
    chart.plot(radarData);
  }

  private static void createScatterPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Scatter scatter,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.ScatterAxes axes = axisRegistry.scatterAxes(scatter.axes());
    XDDFScatterChartData scatterData =
        (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, axes.xAxis(), axes.yAxis());
    scatterData.setVaryColors(scatter.varyColors());
    scatterData.setStyle(toPoiScatterStyle(scatter.style()));
    addSeries(sheet, scatterData, scatter.series(), formulaRuntime);
    chart.plot(scatterData);
  }

  private static void createSurfacePlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Surface surface,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.SurfaceAxes axes = axisRegistry.surfaceAxes(surface.axes());
    XDDFSurfaceChartData surfaceData =
        (XDDFSurfaceChartData)
            chart.createData(ChartTypes.SURFACE, axes.categoryAxis(), axes.valueAxis());
    surfaceData.setVaryColors(surface.varyColors());
    surfaceData.defineSeriesAxis(axes.seriesAxis());
    surfaceData.setWireframe(surface.wireframe());
    addSeries(sheet, surfaceData, surface.series(), formulaRuntime);
    chart.plot(surfaceData);
  }

  private static void createSurface3DPlot(
      XSSFSheet sheet,
      XSSFChart chart,
      AxisRegistry axisRegistry,
      ResolvedChart.Surface3D surface3D,
      @Nullable FormulaRuntime formulaRuntime) {
    AxisRegistry.SurfaceAxes axes = axisRegistry.surfaceAxes(surface3D.axes());
    XDDFSurface3DChartData surfaceData =
        (XDDFSurface3DChartData)
            chart.createData(ChartTypes.SURFACE3D, axes.categoryAxis(), axes.valueAxis());
    surfaceData.setVaryColors(surface3D.varyColors());
    surfaceData.defineSeriesAxis(axes.seriesAxis());
    surfaceData.setWireframe(surface3D.wireframe());
    addSeries(sheet, surfaceData, surface3D.series(), formulaRuntime);
    chart.plot(surfaceData);
  }

  private static void addSeries(
      XSSFSheet sheet,
      XDDFChartData data,
      List<ResolvedChart.Series> definitions,
      @Nullable FormulaRuntime formulaRuntime) {
    for (ResolvedChart.Series definition : definitions) {
      XDDFChartData.Series series =
          data.addSeries(
              toCategoryDataSource(sheet, definition.categories(), formulaRuntime),
              toValueDataSource(sheet, definition.values(), formulaRuntime));
      applySeriesTitle(sheet, series, definition.title(), formulaRuntime);
      applySeriesOptions(series, definition);
    }
  }

  private static void applySeriesTitle(
      XSSFSheet sheet,
      XDDFChartData.Series series,
      ResolvedChartTitle title,
      @Nullable FormulaRuntime formulaRuntime) {
    switch (prepareSeriesTitle(sheet, title, formulaRuntime)) {
      case ChartReport.Title.None _ -> {
        // Leave the series title unset.
      }
      case ChartReport.Title.Text text -> series.setTitle(text.text());
      case ChartReport.Title.Formula formula ->
          series.setTitle(formula.cachedText(), new CellReference(formula.formula()));
    }
  }

  public static void applySeriesOptions(
      XDDFChartData.Series series, ResolvedChart.Series definition) {
    switch (series) {
      case XDDFLineChartData.Series lineSeries -> {
        definition.smooth().ifPresent(lineSeries::setSmooth);
        definition
            .markerStyle()
            .ifPresent(markerStyle -> lineSeries.setMarkerStyle(toPoi(markerStyle)));
        definition.markerSize().ifPresent(lineSeries::setMarkerSize);
      }
      case XDDFLine3DChartData.Series lineSeries -> {
        definition.smooth().ifPresent(lineSeries::setSmooth);
        definition
            .markerStyle()
            .ifPresent(markerStyle -> lineSeries.setMarkerStyle(toPoi(markerStyle)));
        definition.markerSize().ifPresent(lineSeries::setMarkerSize);
      }
      case XDDFScatterChartData.Series scatterSeries -> {
        definition.smooth().ifPresent(scatterSeries::setSmooth);
        definition
            .markerStyle()
            .ifPresent(markerStyle -> scatterSeries.setMarkerStyle(toPoi(markerStyle)));
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

  public static List<ChartReport.Plot> snapshotPlots(
      XSSFChart chart, List<XDDFChartData> chartData) {
    return snapshotPlots(chart, chart.getGraphicFrame(), chartData, null);
  }

  public static List<ChartReport.Plot> snapshotPlots(
      XSSFChart chart, List<XDDFChartData> chartData, @Nullable FormulaRuntime formulaRuntime) {
    return snapshotPlots(chart, chart.getGraphicFrame(), chartData, formulaRuntime);
  }

  public static List<ChartReport.Plot> snapshotPlots(
      XSSFChart chart,
      XSSFGraphicFrame graphicFrame,
      List<XDDFChartData> chartData,
      @Nullable FormulaRuntime formulaRuntime) {
    if (chartData.isEmpty()) {
      return List.of(new ChartReport.Unsupported("UNKNOWN", "Chart contains no parsed plots"));
    }

    List<ChartReport.Plot> plots = new ArrayList<>();
    PlotAreaState state = new PlotAreaState(chart);
    PlotCursor cursor = new PlotCursor();
    XSSFSheet contextSheet = contextSheet(chart, graphicFrame);
    for (XDDFChartData data : chartData) {
      try {
        plots.add(snapshotPlot(contextSheet, state, cursor, data, formulaRuntime));
      } catch (RuntimeException exception) {
        plots.add(
            new ChartReport.Unsupported(
                plotTypeToken(data),
                Objects.requireNonNullElse(
                    exception.getMessage(),
                    "Chart plot family is outside the current modeled chart contract.")));
      }
    }
    return List.copyOf(plots);
  }

  private static ChartReport.Plot snapshotPlot(
      @Nullable XSSFSheet contextSheet,
      PlotAreaState state,
      PlotCursor cursor,
      XDDFChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    return switch (data) {
      case org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData areaData -> {
        CTAreaChart areaChart = state.plotArea().getAreaChartArray(cursor.nextArea());
        yield new ChartReport.Area(
            truthy(areaChart.getVaryColors()),
            fromGroupingTokenOrDefault(
                areaChart.isSetGrouping() ? areaChart.getGrouping().getVal().toString() : null),
            snapshotAxes(state.axesById(), areaChart.getAxIdList()),
            snapshotAreaSeries(contextSheet, areaData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData areaData -> {
        CTArea3DChart areaChart = state.plotArea().getArea3DChartArray(cursor.nextArea3D());
        yield new ChartReport.Area3D(
            truthy(areaChart.getVaryColors()),
            fromGroupingTokenOrDefault(
                areaChart.isSetGrouping() ? areaChart.getGrouping().getVal().toString() : null),
            Optional.ofNullable(areaData.getGapDepth()),
            snapshotAxes(state.axesById(), areaChart.getAxIdList()),
            snapshotArea3DSeries(contextSheet, areaData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFBarChartData barData -> {
        CTBarChart barChart = state.plotArea().getBarChartArray(cursor.nextBar());
        yield new ChartReport.Bar(
            truthy(barChart.getVaryColors()),
            fromPoiBarDirection(barData.getBarDirection()),
            fromBarGroupingTokenOrDefault(
                barChart.isSetGrouping() ? barChart.getGrouping().getVal().toString() : null),
            Optional.ofNullable(barData.getGapWidth()),
            Optional.ofNullable(barData.getOverlap()).map(Integer::valueOf),
            snapshotAxes(state.axesById(), barChart.getAxIdList()),
            snapshotBarSeries(contextSheet, barData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData barData -> {
        CTBar3DChart barChart = state.plotArea().getBar3DChartArray(cursor.nextBar3D());
        yield new ChartReport.Bar3D(
            truthy(barChart.getVaryColors()),
            fromPoiBarDirection(barData.getBarDirection()),
            fromBarGroupingTokenOrDefault(
                barChart.isSetGrouping() ? barChart.getGrouping().getVal().toString() : null),
            Optional.ofNullable(barData.getGapDepth()),
            Optional.ofNullable(barData.getGapWidth()),
            barChart.isSetShape() ? fromPoiBarShape(barData.getShape()) : Optional.empty(),
            snapshotAxes(state.axesById(), barChart.getAxIdList()),
            snapshotBar3DSeries(contextSheet, barData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData doughnutData -> {
        CTDoughnutChart doughnutChart =
            state.plotArea().getDoughnutChartArray(cursor.nextDoughnut());
        yield new ChartReport.Doughnut(
            truthy(doughnutChart.getVaryColors()),
            Optional.ofNullable(doughnutData.getFirstSliceAngle()),
            Optional.ofNullable(doughnutData.getHoleSize()),
            snapshotDoughnutSeries(contextSheet, doughnutData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFLineChartData lineData -> {
        CTLineChart lineChart = state.plotArea().getLineChartArray(cursor.nextLine());
        yield new ChartReport.Line(
            truthy(lineChart.getVaryColors()),
            fromGroupingTokenOrDefault(
                lineChart.getGrouping() == null
                    ? null
                    : lineChart.getGrouping().getVal().toString()),
            snapshotAxes(state.axesById(), lineChart.getAxIdList()),
            snapshotLineSeries(contextSheet, lineData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData lineData -> {
        CTLine3DChart lineChart = state.plotArea().getLine3DChartArray(cursor.nextLine3D());
        yield new ChartReport.Line3D(
            truthy(lineChart.getVaryColors()),
            fromGroupingTokenOrDefault(
                lineChart.getGrouping() == null
                    ? null
                    : lineChart.getGrouping().getVal().toString()),
            Optional.ofNullable(lineData.getGapDepth()),
            snapshotAxes(state.axesById(), lineChart.getAxIdList()),
            snapshotLine3DSeries(contextSheet, lineData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFPieChartData pieData -> {
        CTPieChart pieChart = state.plotArea().getPieChartArray(cursor.nextPie());
        yield new ChartReport.Pie(
            truthy(pieChart.getVaryColors()),
            Optional.ofNullable(pieData.getFirstSliceAngle()),
            snapshotPieSeries(contextSheet, pieData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData pie3DData -> {
        CTPie3DChart pie3DChart = state.plotArea().getPie3DChartArray(cursor.nextPie3D());
        yield new ChartReport.Pie3D(
            truthy(pie3DChart.getVaryColors()),
            snapshotPie3DSeries(contextSheet, pie3DData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData radarData -> {
        CTRadarChart radarChart = state.plotArea().getRadarChartArray(cursor.nextRadar());
        yield new ChartReport.Radar(
            truthy(radarChart.getVaryColors()),
            fromPoiRadarStyle(radarData.getStyle()),
            snapshotAxes(state.axesById(), radarChart.getAxIdList()),
            snapshotRadarSeries(contextSheet, radarData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData scatterData -> {
        CTScatterChart scatterChart = state.plotArea().getScatterChartArray(cursor.nextScatter());
        yield new ChartReport.Scatter(
            truthy(scatterChart.getVaryColors()),
            fromPoiScatterStyle(scatterData.getStyle()),
            snapshotAxes(state.axesById(), scatterChart.getAxIdList()),
            snapshotScatterSeries(contextSheet, scatterData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData surfaceData -> {
        CTSurfaceChart surfaceChart = state.plotArea().getSurfaceChartArray(cursor.nextSurface());
        yield new ChartReport.Surface(
            false,
            Boolean.TRUE.equals(surfaceData.isWireframe()),
            snapshotAxes(state.axesById(), surfaceChart.getAxIdList()),
            snapshotSurfaceSeries(contextSheet, surfaceData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData surface3DData -> {
        CTSurface3DChart surface3DChart =
            state.plotArea().getSurface3DChartArray(cursor.nextSurface3D());
        yield new ChartReport.Surface3D(
            false,
            Boolean.TRUE.equals(surface3DData.isWireframe()),
            snapshotAxes(state.axesById(), surface3DChart.getAxIdList()),
            snapshotSurface3DSeries(contextSheet, surface3DData, formulaRuntime));
      }
      default ->
          new ChartReport.Unsupported(
              plotTypeToken(data),
              "Chart plot family is outside the current modeled chart contract.");
    };
  }

  public static List<ChartReport.Axis> snapshotAxes(
      Map<Long, XDDFChartAxis> axesById, List<CTUnsignedInt> axisIds) {
    List<ChartReport.Axis> axes = new ArrayList<>();
    for (CTUnsignedInt axisId : axisIds) {
      XDDFChartAxis axis = axesById.get(axisId.getVal());
      if (axis != null) {
        axes.add(
            new ChartReport.Axis(
                axisKind(axis),
                fromPoiAxisPosition(axis.getPosition()),
                fromPoiAxisCrosses(axis.getCrosses()),
                axis.isVisible()));
      }
    }
    return List.copyOf(axes);
  }

  private static boolean truthy(CTBoolean value) {
    return value != null && value.getVal();
  }

  /** Immutable view of the parsed plot area and its indexed axes. */
  private record PlotAreaState(CTPlotArea plotArea, Map<Long, XDDFChartAxis> axesById) {
    private PlotAreaState(XSSFChart chart) {
      this(chart.getCTChart().getPlotArea(), indexAxes(chart.getAxes()));
    }
  }

  private static Map<Long, XDDFChartAxis> indexAxes(List<? extends XDDFChartAxis> axes) {
    Map<Long, XDDFChartAxis> axesById = new ConcurrentHashMap<>();
    for (XDDFChartAxis axis : axes) {
      axesById.put(axis.getId(), axis);
    }
    return Map.copyOf(axesById);
  }

  /** Cursor that tracks per-plot-family array indexes while reading the OOXML plot area. */
  private static final class PlotCursor {
    private int area;
    private int area3D;
    private int bar;
    private int bar3D;
    private int doughnut;
    private int line;
    private int line3D;
    private int pie;
    private int pie3D;
    private int radar;
    private int scatter;
    private int surface;
    private int surface3D;

    private int nextArea() {
      int next = area;
      area = area + 1;
      return next;
    }

    private int nextArea3D() {
      int next = area3D;
      area3D = area3D + 1;
      return next;
    }

    private int nextBar() {
      int next = bar;
      bar = bar + 1;
      return next;
    }

    private int nextBar3D() {
      int next = bar3D;
      bar3D = bar3D + 1;
      return next;
    }

    private int nextDoughnut() {
      int next = doughnut;
      doughnut = doughnut + 1;
      return next;
    }

    private int nextLine() {
      int next = line;
      line = line + 1;
      return next;
    }

    private int nextLine3D() {
      int next = line3D;
      line3D = line3D + 1;
      return next;
    }

    private int nextPie() {
      int next = pie;
      pie = pie + 1;
      return next;
    }

    private int nextPie3D() {
      int next = pie3D;
      pie3D = pie3D + 1;
      return next;
    }

    private int nextRadar() {
      int next = radar;
      radar = radar + 1;
      return next;
    }

    private int nextScatter() {
      int next = scatter;
      scatter = scatter + 1;
      return next;
    }

    private int nextSurface() {
      int next = surface;
      surface = surface + 1;
      return next;
    }

    private int nextSurface3D() {
      int next = surface3D;
      surface3D = surface3D + 1;
      return next;
    }
  }

  public static ExcelChartBarDirection fromPoiBarDirection(BarDirection direction) {
    Objects.requireNonNull(direction, "direction must not be null");
    if (direction == BarDirection.COL) {
      return ExcelChartBarDirection.COLUMN;
    }
    return ExcelChartBarDirection.BAR;
  }

  public static BarDirection toPoiBarDirection(ExcelChartBarDirection direction) {
    Objects.requireNonNull(direction, "direction must not be null");
    if (direction == ExcelChartBarDirection.COLUMN) {
      return BarDirection.COL;
    }
    return BarDirection.BAR;
  }

  public static ExcelChartBarGrouping fromPoiBarGroupingOrDefault(BarGrouping grouping) {
    if (grouping == null) {
      return ExcelChartBarGrouping.CLUSTERED;
    }
    return switch (grouping) {
      case STANDARD -> ExcelChartBarGrouping.STANDARD;
      case CLUSTERED -> ExcelChartBarGrouping.CLUSTERED;
      case STACKED -> ExcelChartBarGrouping.STACKED;
      case PERCENT_STACKED -> ExcelChartBarGrouping.PERCENT_STACKED;
    };
  }

  public static ExcelChartBarGrouping fromBarGroupingTokenOrDefault(@Nullable String token) {
    if (token == null) {
      return ExcelChartBarGrouping.CLUSTERED;
    }
    return switch (token.trim().toLowerCase(Locale.ROOT)) {
      case "standard" -> ExcelChartBarGrouping.STANDARD;
      case "clustered" -> ExcelChartBarGrouping.CLUSTERED;
      case "stacked" -> ExcelChartBarGrouping.STACKED;
      case "percentstacked" -> ExcelChartBarGrouping.PERCENT_STACKED;
      default -> throw new IllegalArgumentException("Unsupported bar grouping token: " + token);
    };
  }

  public static BarGrouping toPoiBarGrouping(ExcelChartBarGrouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> BarGrouping.STANDARD;
      case CLUSTERED -> BarGrouping.CLUSTERED;
      case STACKED -> BarGrouping.STACKED;
      case PERCENT_STACKED -> BarGrouping.PERCENT_STACKED;
    };
  }

  public static ExcelChartGrouping fromPoiGrouping(Grouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> ExcelChartGrouping.STANDARD;
      case STACKED -> ExcelChartGrouping.STACKED;
      case PERCENT_STACKED -> ExcelChartGrouping.PERCENT_STACKED;
    };
  }

  public static ExcelChartGrouping fromPoiGroupingOrDefault(Grouping grouping) {
    return grouping == null ? ExcelChartGrouping.STANDARD : fromPoiGrouping(grouping);
  }

  public static ExcelChartGrouping fromGroupingTokenOrDefault(@Nullable String token) {
    if (token == null) {
      return ExcelChartGrouping.STANDARD;
    }
    return switch (token.trim().toLowerCase(Locale.ROOT)) {
      case "standard" -> ExcelChartGrouping.STANDARD;
      case "stacked" -> ExcelChartGrouping.STACKED;
      case "percentstacked" -> ExcelChartGrouping.PERCENT_STACKED;
      default -> throw new IllegalArgumentException("Unsupported grouping token: " + token);
    };
  }

  public static Grouping toPoiGrouping(ExcelChartGrouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> Grouping.STANDARD;
      case STACKED -> Grouping.STACKED;
      case PERCENT_STACKED -> Grouping.PERCENT_STACKED;
    };
  }

  public static Optional<ExcelChartBarShape> fromPoiBarShape(@Nullable Shape shape) {
    if (shape == null) {
      return Optional.empty();
    }
    return Optional.of(
        switch (shape) {
          case BOX -> ExcelChartBarShape.BOX;
          case CONE -> ExcelChartBarShape.CONE;
          case CONE_TO_MAX -> ExcelChartBarShape.CONE_TO_MAX;
          case CYLINDER -> ExcelChartBarShape.CYLINDER;
          case PYRAMID -> ExcelChartBarShape.PYRAMID;
          case PYRAMID_TO_MAX -> ExcelChartBarShape.PYRAMID_TO_MAX;
        });
  }

  public static Shape toPoiBarShape(ExcelChartBarShape shape) {
    Objects.requireNonNull(shape, "shape must not be null");
    return switch (shape) {
      case BOX -> Shape.BOX;
      case CONE -> Shape.CONE;
      case CONE_TO_MAX -> Shape.CONE_TO_MAX;
      case CYLINDER -> Shape.CYLINDER;
      case PYRAMID -> Shape.PYRAMID;
      case PYRAMID_TO_MAX -> Shape.PYRAMID_TO_MAX;
    };
  }

  public static ExcelChartRadarStyle fromPoiRadarStyle(RadarStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case FILLED -> ExcelChartRadarStyle.FILLED;
      case MARKER -> ExcelChartRadarStyle.MARKER;
      case STANDARD -> ExcelChartRadarStyle.STANDARD;
    };
  }

  public static RadarStyle toPoiRadarStyle(ExcelChartRadarStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case FILLED -> RadarStyle.FILLED;
      case MARKER -> RadarStyle.MARKER;
      case STANDARD -> RadarStyle.STANDARD;
    };
  }

  public static ExcelChartScatterStyle fromPoiScatterStyle(ScatterStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case LINE -> ExcelChartScatterStyle.LINE;
      case LINE_MARKER -> ExcelChartScatterStyle.LINE_MARKER;
      case MARKER -> ExcelChartScatterStyle.MARKER;
      case NONE -> ExcelChartScatterStyle.NONE;
      case SMOOTH -> ExcelChartScatterStyle.SMOOTH;
      case SMOOTH_MARKER -> ExcelChartScatterStyle.SMOOTH_MARKER;
    };
  }

  public static ScatterStyle toPoiScatterStyle(ExcelChartScatterStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case LINE -> ScatterStyle.LINE;
      case LINE_MARKER -> ScatterStyle.LINE_MARKER;
      case MARKER -> ScatterStyle.MARKER;
      case NONE -> ScatterStyle.NONE;
      case SMOOTH -> ScatterStyle.SMOOTH;
      case SMOOTH_MARKER -> ScatterStyle.SMOOTH_MARKER;
    };
  }

  public static ExcelChartLegendPosition fromPoiLegendPosition(LegendPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> ExcelChartLegendPosition.BOTTOM;
      case LEFT -> ExcelChartLegendPosition.LEFT;
      case RIGHT -> ExcelChartLegendPosition.RIGHT;
      case TOP -> ExcelChartLegendPosition.TOP;
      case TOP_RIGHT -> ExcelChartLegendPosition.TOP_RIGHT;
    };
  }

  public static LegendPosition toPoiLegendPosition(ExcelChartLegendPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> LegendPosition.BOTTOM;
      case LEFT -> LegendPosition.LEFT;
      case RIGHT -> LegendPosition.RIGHT;
      case TOP -> LegendPosition.TOP;
      case TOP_RIGHT -> LegendPosition.TOP_RIGHT;
    };
  }

  public static ExcelChartDisplayBlanksAs fromPoiDisplayBlanks(STDispBlanksAs.Enum displayBlanks) {
    Objects.requireNonNull(displayBlanks, "displayBlanks must not be null");
    if (displayBlanks == STDispBlanksAs.GAP) {
      return ExcelChartDisplayBlanksAs.GAP;
    }
    if (displayBlanks == STDispBlanksAs.SPAN) {
      return ExcelChartDisplayBlanksAs.SPAN;
    }
    return ExcelChartDisplayBlanksAs.ZERO;
  }

  public static DisplayBlanks toPoiDisplayBlanks(ExcelChartDisplayBlanksAs displayBlanksAs) {
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    return switch (displayBlanksAs) {
      case GAP -> DisplayBlanks.GAP;
      case SPAN -> DisplayBlanks.SPAN;
      case ZERO -> DisplayBlanks.ZERO;
    };
  }

  public static ExcelChartAxisPosition fromPoiAxisPosition(AxisPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> ExcelChartAxisPosition.BOTTOM;
      case LEFT -> ExcelChartAxisPosition.LEFT;
      case RIGHT -> ExcelChartAxisPosition.RIGHT;
      case TOP -> ExcelChartAxisPosition.TOP;
    };
  }

  public static AxisPosition toPoiAxisPosition(ExcelChartAxisPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> AxisPosition.BOTTOM;
      case LEFT -> AxisPosition.LEFT;
      case RIGHT -> AxisPosition.RIGHT;
      case TOP -> AxisPosition.TOP;
    };
  }

  public static ExcelChartAxisCrosses fromPoiAxisCrosses(AxisCrosses crosses) {
    Objects.requireNonNull(crosses, "crosses must not be null");
    return switch (crosses) {
      case AUTO_ZERO -> ExcelChartAxisCrosses.AUTO_ZERO;
      case MAX -> ExcelChartAxisCrosses.MAX;
      case MIN -> ExcelChartAxisCrosses.MIN;
    };
  }

  public static AxisCrosses toPoiAxisCrosses(ExcelChartAxisCrosses crosses) {
    Objects.requireNonNull(crosses, "crosses must not be null");
    return switch (crosses) {
      case AUTO_ZERO -> AxisCrosses.AUTO_ZERO;
      case MAX -> AxisCrosses.MAX;
      case MIN -> AxisCrosses.MIN;
    };
  }

  public static @Nullable XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, @Nullable XSSFGraphicFrame graphicFrame) {
    return optionalChartForGraphicFrame(drawing, graphicFrame).orElse(null);
  }

  public static Optional<XSSFSheet> contextSheet(@Nullable XSSFGraphicFrame graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    return Optional.of(graphicFrame.getDrawing().getSheet());
  }

  public static @Nullable XSSFSheet contextSheet(
      @Nullable XSSFChart chart, @Nullable XSSFGraphicFrame graphicFrame) {
    XSSFGraphicFrame resolvedGraphicFrame =
        graphicFrame != null ? graphicFrame : chart == null ? null : chart.getGraphicFrame();
    return contextSheet(resolvedGraphicFrame).orElse(null);
  }

  public static Optional<String> chartRelationId(XSSFGraphicFrame graphicFrame) {
    return chartRelationNodes(graphicFrame).flatMap(ChartAdapter::chartRelationId);
  }

  public static Optional<String> chartRelationId(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
          graphicFrame) {
    return chartRelationNodes(graphicFrame).flatMap(ChartAdapter::chartRelationId);
  }

  private static Optional<NodeList> chartRelationNodes(@Nullable XSSFGraphicFrame graphicFrame) {
    return chartRelationNodes(
        graphicFrame == null ? null : graphicFrame.getCTGraphicalObjectFrame());
  }

  private static Optional<NodeList> chartRelationNodes(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.@Nullable CTGraphicalObjectFrame
          graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    var graphic = graphicFrame.getGraphic();
    if (graphic == null || graphic.getGraphicData() == null) {
      return Optional.empty();
    }
    return Optional.of(graphic.getGraphicData().getDomNode().getChildNodes());
  }

  public static Optional<String> chartRelationId(Node node) {
    if (!isChartNode(node)) {
      return Optional.empty();
    }
    return relationAttributeValue(node.getAttributes());
  }

  public static Optional<String> relationAttributeValue(NamedNodeMap attributes) {
    if (attributes == null) {
      return Optional.empty();
    }
    Node relationAttribute =
        attributes.getNamedItemNS(
            QName.valueOf("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                .getNamespaceURI(),
            "id");
    if (relationAttribute == null) {
      relationAttribute = attributes.getNamedItem("r:id");
    }
    if (relationAttribute == null || relationAttribute.getNodeValue().isBlank()) {
      return Optional.empty();
    }
    return Optional.of(relationAttribute.getNodeValue());
  }

  public static boolean isChartNode(Node node) {
    return node != null
        && (("chart".equals(node.getLocalName())
                && "http://schemas.openxmlformats.org/drawingml/2006/chart"
                    .equals(node.getNamespaceURI()))
            || "c:chart".equals(node.getNodeName()));
  }

  private static Optional<String> chartRelationId(NodeList nodes) {
    for (int index = 0; index < nodes.getLength(); index++) {
      Optional<String> relationId = chartRelationId(nodes.item(index));
      if (relationId.isPresent()) {
        return relationId;
      }
    }
    return Optional.empty();
  }

  private static Optional<XSSFChart> optionalChartForGraphicFrame(
      XSSFDrawing drawing, @Nullable XSSFGraphicFrame graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    Optional<String> relationId = chartRelationId(graphicFrame);
    if (relationId.isEmpty()) {
      return Optional.empty();
    }
    POIXMLDocumentPart relation = drawing.getRelationById(relationId.orElseThrow());
    return relation instanceof XSSFChart chart ? Optional.of(chart) : Optional.empty();
  }

  private static List<ChartReport.Series> snapshotAreaSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTAreaSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotArea3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTAreaSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotBarSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTBarSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotBar3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTBarSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotDoughnutSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTPieSer().getTx(),
              null,
              null,
              null,
              value.getExplosion(),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotLineSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTLineSer().getTx(),
              smooth(value.getCTLineSer().isSetSmooth(), value.isSmooth()),
              markerStyle(value.getCTLineSer().getMarker()).orElse(null),
              markerSize(value.getCTLineSer().getMarker()),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotLine3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTLineSer().getTx(),
              smooth(value.getCTLineSer().isSetSmooth(), value.isSmooth()),
              markerStyle(value.getCTLineSer().getMarker()).orElse(null),
              markerSize(value.getCTLineSer().getMarker()),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotPieSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTPieSer().getTx(),
              null,
              null,
              null,
              value.getExplosion(),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotPie3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTPieSer().getTx(),
              null,
              null,
              null,
              value.getExplosion(),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotRadarSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTRadarSer().getTx(),
              null,
              null,
              null,
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotScatterSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTScatterSer().getTx(),
              smooth(value.getCTScatterSer().isSetSmooth(), value.isSmooth()),
              markerStyle(value.getCTScatterSer().getMarker()).orElse(null),
              markerSize(value.getCTScatterSer().getMarker()),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotSurfaceSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTSurfaceSer().getTx(),
              null,
              null,
              null,
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static List<ChartReport.Series> snapshotSurface3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData data,
      @Nullable FormulaRuntime formulaRuntime) {
    List<ChartReport.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTSurfaceSer().getTx(),
              null,
              null,
              null,
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static ChartReport.Series snapshotSeries(
      @Nullable XSSFSheet contextSheet,
      XDDFChartData.Series series,
      CTSerTx title,
      @Nullable Boolean smooth,
      @Nullable ExcelChartMarkerStyle markerStyle,
      @Nullable Short markerSize,
      @Nullable FormulaRuntime formulaRuntime) {
    return snapshotSeries(
        contextSheet, series, title, smooth, markerStyle, markerSize, null, formulaRuntime);
  }

  private static ChartReport.Series snapshotSeries(
      @Nullable XSSFSheet contextSheet,
      XDDFChartData.Series series,
      CTSerTx title,
      @Nullable Boolean smooth,
      @Nullable ExcelChartMarkerStyle markerStyle,
      @Nullable Short markerSize,
      @Nullable Long explosion,
      @Nullable FormulaRuntime formulaRuntime) {
    return new ChartReport.Series(
        snapshotSeriesTitle(contextSheet, title, formulaRuntime),
        snapshotDataSource(contextSheet, series.getCategoryData(), formulaRuntime),
        snapshotDataSource(contextSheet, series.getValuesData(), formulaRuntime),
        Optional.ofNullable(smooth),
        Optional.ofNullable(markerStyle),
        Optional.ofNullable(markerSize),
        Optional.ofNullable(explosion));
  }

  private static @Nullable Boolean smooth(boolean present, boolean value) {
    return present ? value : null;
  }

  public static ChartReport.DataSource snapshotDataSource(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      @Nullable FormulaRuntime formulaRuntime) {
    if (source == null) {
      throw new IllegalStateException("Chart series is missing its data source");
    }
    if (source.isReference()) {
      String referenceFormula = source.getDataRangeReference();
      List<String> values =
          resolvedOrCachedReferenceValues(contextSheet, referenceFormula, source, formulaRuntime);
      return source.isNumeric()
          ? new ChartReport.DataSource.NumericReference(
              referenceFormula, Optional.ofNullable(source.getFormatCode()), values)
          : new ChartReport.DataSource.StringReference(referenceFormula, values);
    }
    List<String> values = cachedPointValues(source);
    return source.isNumeric()
        ? new ChartReport.DataSource.NumericLiteral(
            Optional.ofNullable(source.getFormatCode()), values)
        : new ChartReport.DataSource.StringLiteral(values);
  }

  public static List<String> resolvedOrCachedReferenceValues(
      @Nullable XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      @Nullable FormulaRuntime formulaRuntime) {
    if (contextSheet != null && referenceFormula != null && !referenceFormula.isBlank()) {
      try {
        return resolveChartSource(contextSheet, referenceFormula, formulaRuntime).stringValues();
      } catch (RuntimeException ignored) {
        // Fall back to the embedded chart cache when the reference cannot be resolved.
      }
    }
    return cachedPointValues(source);
  }

  private static List<String> cachedPointValues(
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    List<String> values = new ArrayList<>();
    for (int index = 0; index < source.getPointCount(); index++) {
      Object point;
      try {
        point = source.getPointAt(index);
      } catch (IndexOutOfBoundsException exception) {
        point = null;
      }
      values.add(point == null ? "" : point.toString());
    }
    return List.copyOf(values);
  }

  public static Optional<ExcelChartMarkerStyle> markerStyle(CTMarker marker) {
    if (marker == null || !marker.isSetSymbol()) {
      return Optional.empty();
    }
    return fromSymbolToken(marker.getSymbol().getVal().toString());
  }

  public static @Nullable Short markerSize(CTMarker marker) {
    return marker != null && marker.isSetSize() ? (short) marker.getSize().getVal() : null;
  }

  public static DrawingObjectReport.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChartDrawingObject(chart, graphicFrame, null);
  }

  public static DrawingObjectReport.Chart snapshotChartDrawingObject(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      @Nullable FormulaRuntime formulaRuntime) {
    ChartReport snapshot = snapshotChart(chart, graphicFrame, formulaRuntime);
    boolean supported =
        snapshot.plots().stream().noneMatch(ChartReport.Unsupported.class::isInstance);
    return new DrawingObjectReport.Chart(
        resolvedChartName(graphicFrame),
        DrawingAdapter.snapshotAnchor(DrawingAdapter.shapeXml(graphicFrame)),
        supported,
        chartPlotTypeTokens(chart),
        titleSummary(snapshot.title()));
  }

  public static ChartReport snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChart(chart, graphicFrame, null);
  }

  public static ChartReport snapshotChart(
      XSSFChart chart, XSSFGraphicFrame graphicFrame, @Nullable FormulaRuntime formulaRuntime) {
    List<XDDFChartData> chartData = chart.getChartSeries();
    List<ChartReport.Plot> plots = snapshotPlots(chart, graphicFrame, chartData, formulaRuntime);
    return new ChartReport(
        resolvedChartName(graphicFrame),
        DrawingAdapter.snapshotAnchor(DrawingAdapter.shapeXml(graphicFrame)),
        snapshotTitle(chart, graphicFrame, formulaRuntime),
        snapshotLegend(chart),
        snapshotDisplayBlanks(chart),
        chart.isPlotOnlyVisibleCells(),
        plots);
  }

  public static boolean barVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfBarChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getBarChartArray(0).getVaryColors());
  }

  public static boolean lineVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfLineChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getLineChartArray(0).getVaryColors());
  }

  public static boolean pieVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfPieChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getPieChartArray(0).getVaryColors());
  }

  public static List<String> resolvedOrCachedReferenceValues(
      XSSFSheet sheet,
      String formula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    return resolvedOrCachedReferenceValues(sheet, formula, source, null);
  }

  private static List<String> chartPlotTypeTokens(XSSFChart chart) {
    return chartPlotTypeTokens(chart.getChartSeries());
  }

  private static List<String> chartPlotTypeTokens(List<XDDFChartData> chartData) {
    List<String> tokens = new ArrayList<>();
    for (XDDFChartData value : chartData) {
      tokens.add(plotTypeToken(value));
    }
    return List.copyOf(tokens);
  }

  private static ChartReport.Legend snapshotLegend(XSSFChart chart) {
    if (!chart.getCTChart().isSetLegend()) {
      return new ChartReport.Legend.Hidden();
    }
    return new ChartReport.Legend.Visible(
        fromPoiLegendPosition(
            new org.apache.poi.xddf.usermodel.chart.XDDFChartLegend(chart.getCTChart())
                .getPosition()));
  }

  private static ExcelChartDisplayBlanksAs snapshotDisplayBlanks(XSSFChart chart) {
    return chart.getCTChart().isSetDispBlanksAs()
        ? fromPoiDisplayBlanks(chart.getCTChart().getDispBlanksAs().getVal())
        : ExcelChartDisplayBlanksAs.GAP;
  }

  public static Optional<Name> resolveDefinedNameReference(XSSFSheet contextSheet, String formula) {
    if (!formula.matches("^[A-Za-z_][A-Za-z0-9_.]*$")) {
      return Optional.empty();
    }

    int contextSheetIndex = contextSheet.getWorkbook().getSheetIndex(contextSheet);
    Name workbookScopedMatch = null;
    for (Name name : contextSheet.getWorkbook().getAllNames()) {
      if (!formula.equals(name.getNameName())) {
        continue;
      }
      if (name.getSheetIndex() == contextSheetIndex) {
        return Optional.of(name);
      }
      if (name.getSheetIndex() < 0) {
        workbookScopedMatch = name;
      }
    }
    return Optional.ofNullable(workbookScopedMatch);
  }

  public static CellScalar scalarFromFormula(Cell cell) {
    Objects.requireNonNull(cell, "cell must not be null");
    return switch (cell.getCachedFormulaResultType()) {
      case STRING -> new CellScalar(CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC -> new CellScalar(CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new CellScalar(CellScalarKind.STRING, Boolean.toString(cell.getBooleanCellValue()), 0d);
      case BLANK, _NONE -> new CellScalar(CellScalarKind.STRING, "", 0d);
      case ERROR ->
          throw new IllegalArgumentException("Chart source formulas must not cache error values");
      case FORMULA ->
          throw new IllegalArgumentException(
              "Chart source formulas must expose a cached scalar result");
    };
  }

  public static String requiredDefinedNameFormula(Name definedName) {
    Objects.requireNonNull(definedName, "definedName must not be null");
    String refersToFormula = definedName.getRefersToFormula();
    if (refersToFormula == null || refersToFormula.isBlank()) {
      throw new IllegalArgumentException(
          "Defined name '"
              + definedName.getNameName()
              + "' does not resolve to one chart source area");
    }
    return refersToFormula;
  }

  public static org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> toCategoryDataSource(
      XSSFSheet sheet, ChartDataSourceSpec source) {
    return toCategoryDataSource(sheet, source, null);
  }

  public static org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> toCategoryDataSource(
      XSSFSheet sheet, ChartDataSourceSpec source, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return switch (source) {
      case ChartDataSourceSpec.Reference reference -> {
        ResolvedChartSource resolved =
            resolveChartSource(sheet, reference.formula(), formulaRuntime);
        yield resolved.numeric()
            ? org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                resolved.numericValues().toArray(Double[]::new), resolved.referenceFormula())
            : org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                resolved.stringValues().toArray(String[]::new), resolved.referenceFormula());
      }
      case ChartDataSourceSpec.StringLiteral literal ->
          org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
              literal.values().toArray(String[]::new));
      case ChartDataSourceSpec.NumericLiteral literal ->
          org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
              literal.values().toArray(Double[]::new));
    };
  }

  public static org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number>
      toValueDataSource(XSSFSheet sheet, ChartDataSourceSpec source) {
    return toValueDataSource(sheet, source, null);
  }

  public static org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number>
      toValueDataSource(
          XSSFSheet sheet, ChartDataSourceSpec source, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return switch (source) {
      case ChartDataSourceSpec.Reference reference -> {
        ResolvedChartSource resolved =
            resolveChartSource(sheet, reference.formula(), formulaRuntime);
        if (!resolved.numeric()) {
          throw new IllegalArgumentException(
              "Chart value source must resolve to numeric cells: " + reference.formula());
        }
        yield org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
            resolved.numericValues().toArray(Double[]::new), resolved.referenceFormula());
      }
      case ChartDataSourceSpec.StringLiteral _ ->
          throw new IllegalArgumentException("Chart value source must be numeric");
      case ChartDataSourceSpec.NumericLiteral literal ->
          org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
              literal.values().toArray(Double[]::new));
    };
  }

  public static ResolvedChartSource resolveChartSource(XSSFSheet sheet, String formula) {
    return resolveChartSource(sheet, formula, null);
  }

  public static ResolvedChartSource resolveChartSource(
      XSSFSheet sheet, String formula, @Nullable FormulaRuntime formulaRuntime) {
    String normalizedFormula = normalizeFormula(formula);
    ResolvedAreaReference resolved = resolveAreaReference(sheet, normalizedFormula);
    List<String> stringValues = new ArrayList<>();
    List<Double> numericValues = new ArrayList<>();
    boolean numeric = true;
    for (CellReference reference : resolved.areaReference().getAllReferencedCells()) {
      var row = resolved.sheet().getRow(reference.getRow());
      Cell cell = row == null ? null : row.getCell(reference.getCol());
      CellScalar scalar = scalar(cell, formulaRuntime);
      if (scalar.kind() == CellScalarKind.NUMERIC) {
        stringValues.add(Double.toString(scalar.number()));
        numericValues.add(scalar.number());
      } else {
        numeric = false;
        stringValues.add(Objects.requireNonNull(scalar.text(), "scalar text must not be null"));
      }
    }
    return new ResolvedChartSource(
        normalizedFormula,
        resolved.sheet(),
        resolved.areaReference(),
        numeric,
        stringValues,
        numericValues);
  }

  public static ResolvedAreaReference resolveAreaReference(XSSFSheet contextSheet, String formula) {
    Optional<Name> definedName = resolveDefinedNameReference(contextSheet, formula);
    if (definedName.isPresent()) {
      return resolveDefinedName(contextSheet, definedName.orElseThrow());
    }

    AreaReference[] references =
        parseContiguousAreaReferences(
            NamedRangeAdapter.normalizeAreaFormulaForPoi(formula),
            "Chart source formula must resolve to one contiguous area: " + formula);
    if (references.length != 1) {
      throw new IllegalArgumentException(
          "Chart source formula must resolve to one contiguous area: " + formula);
    }
    return new ResolvedAreaReference(
        requireAreaReferenceSheet(contextSheet, references[0], formula), references[0]);
  }

  private static CellReference resolveSingleCellReference(
      XSSFSheet sheet, String formula, String detailLabel) {
    ResolvedAreaReference resolved = resolveAreaReference(sheet, formula);
    if (!resolved.areaReference().isSingleCell()) {
      throw new IllegalArgumentException(
          detailLabel + " must resolve to a single cell: " + formula);
    }
    CellReference firstCell = resolved.areaReference().getFirstCell();
    return new CellReference(
        resolved.sheet().getSheetName(), firstCell.getRow(), firstCell.getCol(), true, true);
  }

  public static String scalarText(XSSFSheet sheet, CellReference reference) {
    return scalarText(sheet, reference, null);
  }

  public static String scalarText(
      XSSFSheet sheet, CellReference reference, @Nullable FormulaRuntime formulaRuntime) {
    XSSFSheet targetSheet =
        reference.getSheetName() == null
            ? sheet
            : requireSheet(
                sheet.getWorkbook(), reference.getSheetName(), reference.formatAsString());
    Cell cell =
        targetSheet.getRow(reference.getRow()) == null
            ? null
            : targetSheet.getRow(reference.getRow()).getCell(reference.getCol());
    CellScalar scalar = scalar(cell, formulaRuntime);
    return scalar.kind() == CellScalarKind.NUMERIC
        ? Double.toString(scalar.number())
        : Objects.requireNonNull(scalar.text(), "scalar text must not be null");
  }

  public static String normalizeFormula(String formula) {
    String normalized = TextValidation.requireNonBlank(formula, "formula");
    return normalized.startsWith("=") ? normalized.substring(1) : normalized;
  }

  public static Optional<String> blankAsOptional(String value) {
    return value == null || value.isBlank() ? Optional.empty() : Optional.of(value);
  }

  private static ResolvedAreaReference resolveDefinedName(
      XSSFSheet contextSheet, Name definedName) {
    String refersToFormula = requiredDefinedNameFormula(definedName);
    AreaReference[] references =
        parseContiguousAreaReferences(
            NamedRangeAdapter.normalizeAreaFormulaForPoi(refersToFormula),
            "Defined name '" + definedName.getNameName() + "' must resolve to one contiguous area");
    if (references.length != 1) {
      throw new IllegalArgumentException(
          "Defined name '" + definedName.getNameName() + "' must resolve to one contiguous area");
    }
    XSSFSheet resolvedSheet =
        references[0].getFirstCell().getSheetName() == null
            ? requireDefinedNameSheet(contextSheet, definedName)
            : requireSheet(
                contextSheet.getWorkbook(),
                references[0].getFirstCell().getSheetName(),
                definedName.getNameName());
    return new ResolvedAreaReference(resolvedSheet, references[0]);
  }

  private static XSSFSheet requireAreaReferenceSheet(
      XSSFSheet contextSheet, AreaReference reference, String formula) {
    String sheetName = reference.getFirstCell().getSheetName();
    return sheetName == null
        ? contextSheet
        : requireSheet(contextSheet.getWorkbook(), sheetName, formula);
  }

  private static XSSFSheet requireDefinedNameSheet(XSSFSheet contextSheet, Name definedName) {
    if (definedName.getSheetIndex() >= 0) {
      return contextSheet.getWorkbook().getSheetAt(definedName.getSheetIndex());
    }
    return contextSheet;
  }

  private static AreaReference[] parseContiguousAreaReferences(
      String formula, String failureMessage) {
    try {
      return AreaReference.generateContiguous(SpreadsheetVersion.EXCEL2007, formula);
    } catch (IllegalArgumentException | IllegalStateException exception) {
      throw new IllegalArgumentException(failureMessage, exception);
    }
  }

  private static XSSFSheet requireSheet(XSSFWorkbook workbook, String sheetName, String formula) {
    XSSFSheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new IllegalArgumentException(
          "Chart source formula '" + formula + "' resolves to missing sheet '" + sheetName + "'");
    }
    return sheet;
  }

  private static CellScalar scalar(@Nullable Cell cell, @Nullable FormulaRuntime formulaRuntime) {
    if (cell == null) {
      return new CellScalar(CellScalarKind.STRING, "", 0d);
    }
    return switch (cell.getCellType()) {
      case STRING -> new CellScalar(CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC -> new CellScalar(CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new CellScalar(CellScalarKind.STRING, Boolean.toString(cell.getBooleanCellValue()), 0d);
      case BLANK, _NONE -> new CellScalar(CellScalarKind.STRING, "", 0d);
      case FORMULA -> scalarFromFormula(cell, formulaRuntime);
      case ERROR ->
          throw new IllegalArgumentException("Chart source cells must not contain error values");
    };
  }

  public static CellScalar scalarFromFormula(Cell cell, @Nullable FormulaRuntime formulaRuntime) {
    Objects.requireNonNull(cell, "cell must not be null");
    if (formulaRuntime == null) {
      return scalarFromFormula(cell);
    }
    CellType evaluatedType = formulaRuntime.evaluateFormulaCell(cell);
    return switch (evaluatedType) {
      case STRING -> new CellScalar(CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC -> new CellScalar(CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new CellScalar(CellScalarKind.STRING, Boolean.toString(cell.getBooleanCellValue()), 0d);
      case BLANK, _NONE -> new CellScalar(CellScalarKind.STRING, "", 0d);
      case ERROR ->
          throw new IllegalArgumentException("Chart source formulas must not cache error values");
      case FORMULA ->
          throw new IllegalArgumentException(
              "Chart source formulas must expose a cached scalar result");
    };
  }

  private static final System.Logger LOGGER = System.getLogger(ChartAdapter.class.getName());

  public static ChartReport.Title snapshotTitle(XSSFChart chart) {
    return snapshotTitle(chart, null);
  }

  public static ChartReport.Title snapshotTitle(
      XSSFChart chart, @Nullable FormulaRuntime formulaRuntime) {
    return snapshotTitle(chart, chart.getGraphicFrame(), formulaRuntime);
  }

  public static ChartReport.Title snapshotTitle(
      XSSFChart chart,
      @Nullable XSSFGraphicFrame graphicFrame,
      @Nullable FormulaRuntime formulaRuntime) {
    if (!chart.getCTChart().isSetTitle()) {
      return new ChartReport.Title.None();
    }
    String formula = chart.getTitleFormula();
    if (formula != null) {
      return new ChartReport.Title.Formula(
          formula, cachedTitleText(chart, graphicFrame, formula, formulaRuntime));
    }
    String text = chart.getTitleText().getString();
    return text.isBlank() ? new ChartReport.Title.None() : new ChartReport.Title.Text(text);
  }

  public static String cachedTitleText(XSSFChart chart, String formula) {
    return cachedTitleText(chart, formula, null);
  }

  public static String cachedTitleText(
      XSSFChart chart, String formula, @Nullable FormulaRuntime formulaRuntime) {
    return cachedTitleText(chart, chart.getGraphicFrame(), formula, formulaRuntime);
  }

  public static String cachedTitleText(
      XSSFChart chart,
      @Nullable XSSFGraphicFrame graphicFrame,
      String formula,
      @Nullable FormulaRuntime formulaRuntime) {
    Optional<String> resolvedText =
        optionalResolvedTitleFormulaText(chart, graphicFrame, formula, formulaRuntime);
    if (resolvedText.isPresent()) {
      return resolvedText.orElseThrow();
    }
    if (!chart.getCTChart().isSetTitle()
        || !chart.getCTChart().getTitle().isSetTx()
        || !chart.getCTChart().getTitle().getTx().isSetStrRef()
        || !chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache()
        || chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().sizeOfPtArray() == 0) {
      return "";
    }
    return chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV();
  }

  public static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    return resolvedTitleFormulaText(chart, formula, null);
  }

  public static String resolvedTitleFormulaText(
      XSSFChart chart, String formula, @Nullable FormulaRuntime formulaRuntime) {
    return optionalResolvedTitleFormulaText(chart, chart.getGraphicFrame(), formula, formulaRuntime)
        .orElse("");
  }

  public static Optional<String> optionalResolvedTitleFormulaText(
      XSSFChart chart,
      @Nullable XSSFGraphicFrame graphicFrame,
      String formula,
      @Nullable FormulaRuntime formulaRuntime) {
    try {
      XSSFSheet contextSheet = contextSheet(chart, graphicFrame);
      return contextSheet == null
          ? Optional.empty()
          : Optional.of(
              scalarText(
                  contextSheet,
                  resolveSingleCellReference(contextSheet, formula, "Chart title formula"),
                  formulaRuntime));
    } catch (IllegalArgumentException exception) {
      if (recoverableTitleFormulaResolutionFailure(exception)) {
        return Optional.empty();
      }
      LOGGER.log(
          System.Logger.Level.WARNING,
          "Failed to resolve chart title formula '" + formula + "'; using cached or empty title",
          exception);
      return Optional.empty();
    }
  }

  public static ChartReport.Title snapshotSeriesTitle(CTSerTx title) {
    return snapshotSeriesTitle(null, title, null);
  }

  public static ChartReport.Title snapshotSeriesTitle(
      @Nullable XSSFSheet contextSheet,
      @Nullable CTSerTx title,
      @Nullable FormulaRuntime formulaRuntime) {
    if (title == null) {
      return new ChartReport.Title.None();
    }
    if (title.isSetStrRef()) {
      if (contextSheet != null) {
        try {
          return new ChartReport.Title.Formula(
              title.getStrRef().getF(),
              scalarText(
                  contextSheet,
                  resolveSingleCellReference(
                      contextSheet, title.getStrRef().getF(), "Series title formula"),
                  formulaRuntime));
        } catch (IllegalArgumentException ignored) {
          // Fall back to the embedded chart cache when the formula cannot be resolved live.
        }
      }
      String cachedText =
          title.getStrRef().isSetStrCache() && title.getStrRef().getStrCache().sizeOfPtArray() > 0
              ? title.getStrRef().getStrCache().getPtArray(0).getV()
              : "";
      return new ChartReport.Title.Formula(title.getStrRef().getF(), cachedText);
    }
    return title.isSetV() ? new ChartReport.Title.Text(title.getV()) : new ChartReport.Title.None();
  }

  public static String titleSummary(ChartReport.Title title) {
    return switch (title) {
      case ChartReport.Title.None _ -> "";
      case ChartReport.Title.Text text -> text.text();
      case ChartReport.Title.Formula formula ->
          formula.cachedText().isEmpty() ? formula.formula() : formula.cachedText();
    };
  }

  public static String resolvedChartName(XSSFGraphicFrame graphicFrame) {
    String name = blankAsOptional(graphicFrame.getName()).orElse(null);
    return name != null ? name : "Chart-" + graphicFrame.getId();
  }

  private static boolean recoverableTitleFormulaResolutionFailure(
      IllegalArgumentException exception) {
    return "Chart source formulas must not cache error values".equals(exception.getMessage());
  }

  public static final class AxisRegistry {
    private final XDDFChart chart;
    private final Map<ChartAxisSpec, XDDFChartAxis> axesByDefinition = new ConcurrentHashMap<>();

    public AxisRegistry(XDDFChart chart) {
      this.chart = Objects.requireNonNull(chart, "chart must not be null");
    }

    public CategoryValueAxes categoryValueAxes(List<ChartAxisSpec> definitions) {
      XDDFChartAxis categoryAxis = null;
      XDDFValueAxis valueAxis = null;
      for (ChartAxisSpec definition : definitions) {
        XDDFChartAxis axis = axis(definition);
        if (definition.kind().isCategoryAxis()) {
          categoryAxis = axis;
        }
        if (definition.kind().isValueAxis()) {
          valueAxis = (XDDFValueAxis) axis;
        }
        if (definition.kind().isSeriesAxis()) {
          throw new IllegalArgumentException("Series axis is unsupported for this plot family");
        }
      }
      if (categoryAxis == null || valueAxis == null) {
        throw new IllegalArgumentException(
            "Plot must declare one category/date axis and one value axis");
      }
      crossAxes(List.of(categoryAxis, valueAxis));
      return new CategoryValueAxes(categoryAxis, valueAxis);
    }

    public ScatterAxes scatterAxes(List<ChartAxisSpec> definitions) {
      List<XDDFValueAxis> valueAxes =
          definitions.stream().map(this::axis).map(axis -> (XDDFValueAxis) axis).toList();
      if (valueAxes.size() != 2) {
        throw new IllegalArgumentException("Scatter plots must declare exactly two value axes");
      }
      crossAxes(List.copyOf(valueAxes));
      return new ScatterAxes(valueAxes.getFirst(), valueAxes.get(1));
    }

    public SurfaceAxes surfaceAxes(List<ChartAxisSpec> definitions) {
      XDDFChartAxis categoryAxis = null;
      XDDFValueAxis valueAxis = null;
      XDDFSeriesAxis seriesAxis = null;
      for (ChartAxisSpec definition : definitions) {
        XDDFChartAxis axis = axis(definition);
        if (definition.kind().isCategoryAxis()) {
          categoryAxis = axis;
        }
        if (definition.kind().isValueAxis()) {
          valueAxis = (XDDFValueAxis) axis;
        }
        if (definition.kind().isSeriesAxis()) {
          seriesAxis = (XDDFSeriesAxis) axis;
        }
      }
      if (categoryAxis == null || valueAxis == null || seriesAxis == null) {
        throw new IllegalArgumentException(
            "Surface plots must declare one category/date axis, one value axis, and one series axis");
      }
      crossAxes(List.of(categoryAxis, valueAxis, seriesAxis));
      return new SurfaceAxes(categoryAxis, valueAxis, seriesAxis);
    }

    private XDDFChartAxis axis(ChartAxisSpec definition) {
      XDDFChartAxis existing = axesByDefinition.get(definition);
      if (existing != null) {
        return existing;
      }
      XDDFChartAxis created = createAxis(chart, definition.kind(), definition.position());
      created.setCrosses(toPoiAxisCrosses(definition.crosses()));
      created.setVisible(definition.visible());
      axesByDefinition.put(definition, created);
      return created;
    }

    private static void crossAxes(List<? extends XDDFChartAxis> axes) {
      for (int index = 0; index < axes.size(); index++) {
        for (int other = index + 1; other < axes.size(); other++) {
          axes.get(index).crossAxis(axes.get(other));
          axes.get(other).crossAxis(axes.get(index));
        }
      }
    }

    /** Pair of category/date axis and value axis for standard chart families. */
    public record CategoryValueAxes(XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis) {}

    /** Pair of value axes for scatter charts. */
    public record ScatterAxes(XDDFValueAxis xAxis, XDDFValueAxis yAxis) {}

    /** Triple of category/date, value, and series axes for surface charts. */
    public record SurfaceAxes(
        XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis, XDDFSeriesAxis seriesAxis) {}
  }

  public record ResolvedAreaReference(XSSFSheet sheet, AreaReference areaReference) {
    public ResolvedAreaReference {
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(areaReference, "areaReference must not be null");
    }
  }

  public record ResolvedChartSource(
      String referenceFormula,
      XSSFSheet sheet,
      AreaReference areaReference,
      boolean numeric,
      List<String> stringValues,
      List<Double> numericValues) {
    public ResolvedChartSource {
      Objects.requireNonNull(referenceFormula, "referenceFormula must not be null");
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(areaReference, "areaReference must not be null");
      stringValues =
          List.copyOf(Objects.requireNonNull(stringValues, "stringValues must not be null"));
      numericValues =
          List.copyOf(Objects.requireNonNull(numericValues, "numericValues must not be null"));
    }
  }

  /** Supported scalar kinds extracted from chart source cells. */
  public enum CellScalarKind {
    STRING,
    NUMERIC
  }

  public record CellScalar(CellScalarKind kind, @Nullable String text, double number) {
    public CellScalar {
      Objects.requireNonNull(kind, "kind must not be null");
      if (kind == CellScalarKind.STRING) {
        Objects.requireNonNull(text, "text must not be null");
      }
    }
  }
}
