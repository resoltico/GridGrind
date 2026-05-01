package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTArea3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBar3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLine3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPie3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScatterChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurface3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurfaceChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;

/** Converts POI chart plot parts into immutable GridGrind snapshot models. */
final class ExcelChartPlotSnapshotSupport {
  private ExcelChartPlotSnapshotSupport() {}

  static List<ExcelChartSnapshot.Plot> snapshotPlots(
      XSSFChart chart, List<XDDFChartData> chartData) {
    return snapshotPlots(chart, chart.getGraphicFrame(), chartData, null);
  }

  static List<ExcelChartSnapshot.Plot> snapshotPlots(
      XSSFChart chart, List<XDDFChartData> chartData, ExcelFormulaRuntime formulaRuntime) {
    return snapshotPlots(chart, chart.getGraphicFrame(), chartData, formulaRuntime);
  }

  static List<ExcelChartSnapshot.Plot> snapshotPlots(
      XSSFChart chart,
      XSSFGraphicFrame graphicFrame,
      List<XDDFChartData> chartData,
      ExcelFormulaRuntime formulaRuntime) {
    if (chartData.isEmpty()) {
      return List.of(
          new ExcelChartSnapshot.Unsupported("UNKNOWN", "Chart contains no parsed plots"));
    }

    List<ExcelChartSnapshot.Plot> plots = new ArrayList<>();
    PlotAreaState state = new PlotAreaState(chart);
    PlotCursor cursor = new PlotCursor();
    XSSFSheet contextSheet = ExcelChartSnapshotSupport.contextSheet(chart, graphicFrame);
    for (XDDFChartData data : chartData) {
      try {
        plots.add(snapshotPlot(contextSheet, state, cursor, data, formulaRuntime));
      } catch (RuntimeException exception) {
        plots.add(
            new ExcelChartSnapshot.Unsupported(
                ExcelChartPoiBridge.plotTypeToken(data), exception.getMessage()));
      }
    }
    return List.copyOf(plots);
  }

  private static ExcelChartSnapshot.Plot snapshotPlot(
      XSSFSheet contextSheet,
      PlotAreaState state,
      PlotCursor cursor,
      XDDFChartData data,
      ExcelFormulaRuntime formulaRuntime) {
    return switch (data) {
      case org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData areaData -> {
        CTAreaChart areaChart = state.plotArea().getAreaChartArray(cursor.nextArea());
        yield new ExcelChartSnapshot.Area(
            truthy(areaChart.getVaryColors()),
            ExcelChartPoiBridge.fromGroupingTokenOrDefault(
                areaChart.isSetGrouping() ? areaChart.getGrouping().getVal().toString() : null),
            snapshotAxes(state.axesById(), areaChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotAreaSeries(
                contextSheet, areaData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData areaData -> {
        CTArea3DChart areaChart = state.plotArea().getArea3DChartArray(cursor.nextArea3D());
        yield new ExcelChartSnapshot.Area3D(
            truthy(areaChart.getVaryColors()),
            ExcelChartPoiBridge.fromGroupingTokenOrDefault(
                areaChart.isSetGrouping() ? areaChart.getGrouping().getVal().toString() : null),
            areaData.getGapDepth(),
            snapshotAxes(state.axesById(), areaChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotArea3DSeries(
                contextSheet, areaData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFBarChartData barData -> {
        CTBarChart barChart = state.plotArea().getBarChartArray(cursor.nextBar());
        yield new ExcelChartSnapshot.Bar(
            truthy(barChart.getVaryColors()),
            ExcelChartPoiBridge.fromPoiBarDirection(barData.getBarDirection()),
            ExcelChartPoiBridge.fromBarGroupingTokenOrDefault(
                barChart.isSetGrouping() ? barChart.getGrouping().getVal().toString() : null),
            barData.getGapWidth(),
            barData.getOverlap() == null ? null : Integer.valueOf(barData.getOverlap()),
            snapshotAxes(state.axesById(), barChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotBarSeries(
                contextSheet, barData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData barData -> {
        CTBar3DChart barChart = state.plotArea().getBar3DChartArray(cursor.nextBar3D());
        yield new ExcelChartSnapshot.Bar3D(
            truthy(barChart.getVaryColors()),
            ExcelChartPoiBridge.fromPoiBarDirection(barData.getBarDirection()),
            ExcelChartPoiBridge.fromBarGroupingTokenOrDefault(
                barChart.isSetGrouping() ? barChart.getGrouping().getVal().toString() : null),
            barData.getGapDepth(),
            barData.getGapWidth(),
            barChart.isSetShape() ? ExcelChartPoiBridge.fromPoiBarShape(barData.getShape()) : null,
            snapshotAxes(state.axesById(), barChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotBar3DSeries(
                contextSheet, barData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData doughnutData -> {
        CTDoughnutChart doughnutChart =
            state.plotArea().getDoughnutChartArray(cursor.nextDoughnut());
        yield new ExcelChartSnapshot.Doughnut(
            truthy(doughnutChart.getVaryColors()),
            doughnutData.getFirstSliceAngle(),
            doughnutData.getHoleSize(),
            ExcelChartSeriesSnapshotSupport.snapshotDoughnutSeries(
                contextSheet, doughnutData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFLineChartData lineData -> {
        CTLineChart lineChart = state.plotArea().getLineChartArray(cursor.nextLine());
        yield new ExcelChartSnapshot.Line(
            truthy(lineChart.getVaryColors()),
            ExcelChartPoiBridge.fromGroupingTokenOrDefault(
                lineChart.getGrouping() == null
                    ? null
                    : lineChart.getGrouping().getVal().toString()),
            snapshotAxes(state.axesById(), lineChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotLineSeries(
                contextSheet, lineData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData lineData -> {
        CTLine3DChart lineChart = state.plotArea().getLine3DChartArray(cursor.nextLine3D());
        yield new ExcelChartSnapshot.Line3D(
            truthy(lineChart.getVaryColors()),
            ExcelChartPoiBridge.fromGroupingTokenOrDefault(
                lineChart.getGrouping() == null
                    ? null
                    : lineChart.getGrouping().getVal().toString()),
            lineData.getGapDepth(),
            snapshotAxes(state.axesById(), lineChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotLine3DSeries(
                contextSheet, lineData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFPieChartData pieData -> {
        CTPieChart pieChart = state.plotArea().getPieChartArray(cursor.nextPie());
        yield new ExcelChartSnapshot.Pie(
            truthy(pieChart.getVaryColors()),
            pieData.getFirstSliceAngle(),
            ExcelChartSeriesSnapshotSupport.snapshotPieSeries(
                contextSheet, pieData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData pie3DData -> {
        CTPie3DChart pie3DChart = state.plotArea().getPie3DChartArray(cursor.nextPie3D());
        yield new ExcelChartSnapshot.Pie3D(
            truthy(pie3DChart.getVaryColors()),
            ExcelChartSeriesSnapshotSupport.snapshotPie3DSeries(
                contextSheet, pie3DData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData radarData -> {
        CTRadarChart radarChart = state.plotArea().getRadarChartArray(cursor.nextRadar());
        yield new ExcelChartSnapshot.Radar(
            truthy(radarChart.getVaryColors()),
            ExcelChartPoiBridge.fromPoiRadarStyle(radarData.getStyle()),
            snapshotAxes(state.axesById(), radarChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotRadarSeries(
                contextSheet, radarData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData scatterData -> {
        CTScatterChart scatterChart = state.plotArea().getScatterChartArray(cursor.nextScatter());
        yield new ExcelChartSnapshot.Scatter(
            truthy(scatterChart.getVaryColors()),
            ExcelChartPoiBridge.fromPoiScatterStyle(scatterData.getStyle()),
            snapshotAxes(state.axesById(), scatterChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotScatterSeries(
                contextSheet, scatterData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData surfaceData -> {
        CTSurfaceChart surfaceChart = state.plotArea().getSurfaceChartArray(cursor.nextSurface());
        yield new ExcelChartSnapshot.Surface(
            false,
            Boolean.TRUE.equals(surfaceData.isWireframe()),
            snapshotAxes(state.axesById(), surfaceChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotSurfaceSeries(
                contextSheet, surfaceData, formulaRuntime));
      }
      case org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData surface3DData -> {
        CTSurface3DChart surface3DChart =
            state.plotArea().getSurface3DChartArray(cursor.nextSurface3D());
        yield new ExcelChartSnapshot.Surface3D(
            false,
            Boolean.TRUE.equals(surface3DData.isWireframe()),
            snapshotAxes(state.axesById(), surface3DChart.getAxIdList()),
            ExcelChartSeriesSnapshotSupport.snapshotSurface3DSeries(
                contextSheet, surface3DData, formulaRuntime));
      }
      default ->
          new ExcelChartSnapshot.Unsupported(
              ExcelChartPoiBridge.plotTypeToken(data),
              "Chart plot family is outside the current modeled chart contract.");
    };
  }

  static List<ExcelChartSnapshot.Axis> snapshotAxes(
      Map<Long, XDDFChartAxis> axesById, List<CTUnsignedInt> axisIds) {
    List<ExcelChartSnapshot.Axis> axes = new ArrayList<>();
    for (CTUnsignedInt axisId : axisIds) {
      XDDFChartAxis axis = axesById.get(axisId.getVal());
      if (axis != null) {
        axes.add(
            new ExcelChartSnapshot.Axis(
                ExcelChartPoiBridge.axisKind(axis),
                ExcelChartPoiBridge.fromPoiAxisPosition(axis.getPosition()),
                ExcelChartPoiBridge.fromPoiAxisCrosses(axis.getCrosses()),
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
}
