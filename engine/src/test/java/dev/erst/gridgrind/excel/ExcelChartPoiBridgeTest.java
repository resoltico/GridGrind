package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import java.io.IOException;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.text.XDDFRunProperties;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCrosses;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumFmt;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTickLblPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTickMark;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;

/** Unit tests for the package-owned POI chart translation seam. */
class ExcelChartPoiBridgeTest {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void convertsAllModeledPoiChartEnumsAndTokens() {
    assertEquals(
        ExcelChartBarDirection.COLUMN, ExcelChartPoiBridge.fromPoiBarDirection(BarDirection.COL));
    assertEquals(
        ExcelChartBarDirection.BAR, ExcelChartPoiBridge.fromPoiBarDirection(BarDirection.BAR));
    assertEquals(
        BarDirection.COL, ExcelChartPoiBridge.toPoiBarDirection(ExcelChartBarDirection.COLUMN));
    assertEquals(
        BarDirection.BAR, ExcelChartPoiBridge.toPoiBarDirection(ExcelChartBarDirection.BAR));

    assertEquals(
        ExcelChartBarGrouping.STANDARD,
        ExcelChartPoiBridge.fromPoiBarGrouping(BarGrouping.STANDARD));
    assertEquals(
        ExcelChartBarGrouping.CLUSTERED,
        ExcelChartPoiBridge.fromPoiBarGrouping(BarGrouping.CLUSTERED));
    assertEquals(
        ExcelChartBarGrouping.STACKED, ExcelChartPoiBridge.fromPoiBarGrouping(BarGrouping.STACKED));
    assertEquals(
        ExcelChartBarGrouping.PERCENT_STACKED,
        ExcelChartPoiBridge.fromPoiBarGrouping(BarGrouping.PERCENT_STACKED));
    assertEquals(
        ExcelChartBarGrouping.CLUSTERED, ExcelChartPoiBridge.fromPoiBarGroupingOrDefault(null));
    assertEquals(
        ExcelChartBarGrouping.STANDARD,
        ExcelChartPoiBridge.fromPoiBarGroupingOrDefault(BarGrouping.STANDARD));
    assertEquals(
        ExcelChartBarGrouping.STANDARD,
        ExcelChartPoiBridge.fromBarGroupingTokenOrDefault("standard"));
    assertEquals(
        ExcelChartBarGrouping.CLUSTERED,
        ExcelChartPoiBridge.fromBarGroupingTokenOrDefault("clustered"));
    assertEquals(
        ExcelChartBarGrouping.STACKED,
        ExcelChartPoiBridge.fromBarGroupingTokenOrDefault("stacked"));
    assertEquals(
        ExcelChartBarGrouping.PERCENT_STACKED,
        ExcelChartPoiBridge.fromBarGroupingTokenOrDefault("percentstacked"));
    assertEquals(
        ExcelChartBarGrouping.CLUSTERED, ExcelChartPoiBridge.fromBarGroupingTokenOrDefault(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelChartPoiBridge.fromBarGroupingTokenOrDefault("unsupported"));
    assertEquals(
        BarGrouping.STANDARD, ExcelChartPoiBridge.toPoiBarGrouping(ExcelChartBarGrouping.STANDARD));
    assertEquals(
        BarGrouping.CLUSTERED,
        ExcelChartPoiBridge.toPoiBarGrouping(ExcelChartBarGrouping.CLUSTERED));
    assertEquals(
        BarGrouping.STACKED, ExcelChartPoiBridge.toPoiBarGrouping(ExcelChartBarGrouping.STACKED));
    assertEquals(
        BarGrouping.PERCENT_STACKED,
        ExcelChartPoiBridge.toPoiBarGrouping(ExcelChartBarGrouping.PERCENT_STACKED));

    assertEquals(
        ExcelChartGrouping.STANDARD, ExcelChartPoiBridge.fromPoiGrouping(Grouping.STANDARD));
    assertEquals(ExcelChartGrouping.STACKED, ExcelChartPoiBridge.fromPoiGrouping(Grouping.STACKED));
    assertEquals(
        ExcelChartGrouping.PERCENT_STACKED,
        ExcelChartPoiBridge.fromPoiGrouping(Grouping.PERCENT_STACKED));
    assertEquals(ExcelChartGrouping.STANDARD, ExcelChartPoiBridge.fromPoiGroupingOrDefault(null));
    assertEquals(
        ExcelChartGrouping.STANDARD,
        ExcelChartPoiBridge.fromPoiGroupingOrDefault(Grouping.STANDARD));
    assertEquals(
        ExcelChartGrouping.STANDARD, ExcelChartPoiBridge.fromGroupingTokenOrDefault("standard"));
    assertEquals(
        ExcelChartGrouping.STACKED, ExcelChartPoiBridge.fromGroupingTokenOrDefault("stacked"));
    assertEquals(
        ExcelChartGrouping.PERCENT_STACKED,
        ExcelChartPoiBridge.fromGroupingTokenOrDefault("percentstacked"));
    assertEquals(ExcelChartGrouping.STANDARD, ExcelChartPoiBridge.fromGroupingTokenOrDefault(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelChartPoiBridge.fromGroupingTokenOrDefault("unsupported"));
    assertEquals(Grouping.STANDARD, ExcelChartPoiBridge.toPoiGrouping(ExcelChartGrouping.STANDARD));
    assertEquals(Grouping.STACKED, ExcelChartPoiBridge.toPoiGrouping(ExcelChartGrouping.STACKED));
    assertEquals(
        Grouping.PERCENT_STACKED,
        ExcelChartPoiBridge.toPoiGrouping(ExcelChartGrouping.PERCENT_STACKED));

    assertEquals(ExcelChartBarShape.BOX, ExcelChartPoiBridge.fromPoiBarShape(Shape.BOX));
    assertEquals(ExcelChartBarShape.CONE, ExcelChartPoiBridge.fromPoiBarShape(Shape.CONE));
    assertEquals(
        ExcelChartBarShape.CONE_TO_MAX, ExcelChartPoiBridge.fromPoiBarShape(Shape.CONE_TO_MAX));
    assertEquals(ExcelChartBarShape.CYLINDER, ExcelChartPoiBridge.fromPoiBarShape(Shape.CYLINDER));
    assertEquals(ExcelChartBarShape.PYRAMID, ExcelChartPoiBridge.fromPoiBarShape(Shape.PYRAMID));
    assertEquals(
        ExcelChartBarShape.PYRAMID_TO_MAX,
        ExcelChartPoiBridge.fromPoiBarShape(Shape.PYRAMID_TO_MAX));
    assertEquals(Shape.BOX, ExcelChartPoiBridge.toPoiBarShape(ExcelChartBarShape.BOX));
    assertEquals(Shape.CONE, ExcelChartPoiBridge.toPoiBarShape(ExcelChartBarShape.CONE));
    assertEquals(
        Shape.CONE_TO_MAX, ExcelChartPoiBridge.toPoiBarShape(ExcelChartBarShape.CONE_TO_MAX));
    assertEquals(Shape.CYLINDER, ExcelChartPoiBridge.toPoiBarShape(ExcelChartBarShape.CYLINDER));
    assertEquals(Shape.PYRAMID, ExcelChartPoiBridge.toPoiBarShape(ExcelChartBarShape.PYRAMID));
    assertEquals(
        Shape.PYRAMID_TO_MAX, ExcelChartPoiBridge.toPoiBarShape(ExcelChartBarShape.PYRAMID_TO_MAX));

    assertEquals(
        ExcelChartRadarStyle.FILLED, ExcelChartPoiBridge.fromPoiRadarStyle(RadarStyle.FILLED));
    assertEquals(
        ExcelChartRadarStyle.MARKER, ExcelChartPoiBridge.fromPoiRadarStyle(RadarStyle.MARKER));
    assertEquals(
        ExcelChartRadarStyle.STANDARD, ExcelChartPoiBridge.fromPoiRadarStyle(RadarStyle.STANDARD));
    assertEquals(
        RadarStyle.FILLED, ExcelChartPoiBridge.toPoiRadarStyle(ExcelChartRadarStyle.FILLED));
    assertEquals(
        RadarStyle.MARKER, ExcelChartPoiBridge.toPoiRadarStyle(ExcelChartRadarStyle.MARKER));
    assertEquals(
        RadarStyle.STANDARD, ExcelChartPoiBridge.toPoiRadarStyle(ExcelChartRadarStyle.STANDARD));

    assertEquals(
        ExcelChartScatterStyle.LINE, ExcelChartPoiBridge.fromPoiScatterStyle(ScatterStyle.LINE));
    assertEquals(
        ExcelChartScatterStyle.LINE_MARKER,
        ExcelChartPoiBridge.fromPoiScatterStyle(ScatterStyle.LINE_MARKER));
    assertEquals(
        ExcelChartScatterStyle.MARKER,
        ExcelChartPoiBridge.fromPoiScatterStyle(ScatterStyle.MARKER));
    assertEquals(
        ExcelChartScatterStyle.NONE, ExcelChartPoiBridge.fromPoiScatterStyle(ScatterStyle.NONE));
    assertEquals(
        ExcelChartScatterStyle.SMOOTH,
        ExcelChartPoiBridge.fromPoiScatterStyle(ScatterStyle.SMOOTH));
    assertEquals(
        ExcelChartScatterStyle.SMOOTH_MARKER,
        ExcelChartPoiBridge.fromPoiScatterStyle(ScatterStyle.SMOOTH_MARKER));
    assertEquals(
        ScatterStyle.LINE, ExcelChartPoiBridge.toPoiScatterStyle(ExcelChartScatterStyle.LINE));
    assertEquals(
        ScatterStyle.LINE_MARKER,
        ExcelChartPoiBridge.toPoiScatterStyle(ExcelChartScatterStyle.LINE_MARKER));
    assertEquals(
        ScatterStyle.MARKER, ExcelChartPoiBridge.toPoiScatterStyle(ExcelChartScatterStyle.MARKER));
    assertEquals(
        ScatterStyle.NONE, ExcelChartPoiBridge.toPoiScatterStyle(ExcelChartScatterStyle.NONE));
    assertEquals(
        ScatterStyle.SMOOTH, ExcelChartPoiBridge.toPoiScatterStyle(ExcelChartScatterStyle.SMOOTH));
    assertEquals(
        ScatterStyle.SMOOTH_MARKER,
        ExcelChartPoiBridge.toPoiScatterStyle(ExcelChartScatterStyle.SMOOTH_MARKER));

    assertEquals(
        ExcelChartMarkerStyle.CIRCLE, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.CIRCLE));
    assertEquals(
        ExcelChartMarkerStyle.DASH, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.DASH));
    assertEquals(
        ExcelChartMarkerStyle.DIAMOND, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.DIAMOND));
    assertEquals(
        ExcelChartMarkerStyle.DOT, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.DOT));
    assertEquals(
        ExcelChartMarkerStyle.NONE, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.NONE));
    assertEquals(
        ExcelChartMarkerStyle.PICTURE, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.PICTURE));
    assertEquals(
        ExcelChartMarkerStyle.PLUS, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.PLUS));
    assertEquals(
        ExcelChartMarkerStyle.SQUARE, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.SQUARE));
    assertEquals(
        ExcelChartMarkerStyle.STAR, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.STAR));
    assertEquals(
        ExcelChartMarkerStyle.TRIANGLE,
        ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.TRIANGLE));
    assertEquals(ExcelChartMarkerStyle.X, ExcelChartPoiBridge.fromPoiMarkerStyle(MarkerStyle.X));
    for (ExcelChartMarkerStyle style : ExcelChartMarkerStyle.values()) {
      assertEquals(
          style,
          ExcelChartPoiBridge.fromPoiMarkerStyle(ExcelChartPoiBridge.toPoiMarkerStyle(style)));
    }

    assertEquals(
        ExcelChartLegendPosition.BOTTOM,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.BOTTOM));
    assertEquals(
        ExcelChartLegendPosition.LEFT,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.LEFT));
    assertEquals(
        ExcelChartLegendPosition.RIGHT,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.RIGHT));
    assertEquals(
        ExcelChartLegendPosition.TOP,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.TOP));
    assertEquals(
        ExcelChartLegendPosition.TOP_RIGHT,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.TOP_RIGHT));
    assertEquals(
        LegendPosition.BOTTOM,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.BOTTOM));
    assertEquals(
        LegendPosition.LEFT,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.LEFT));
    assertEquals(
        LegendPosition.RIGHT,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.RIGHT));
    assertEquals(
        LegendPosition.TOP, ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.TOP));
    assertEquals(
        LegendPosition.TOP_RIGHT,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.TOP_RIGHT));

    assertEquals(
        ExcelChartDisplayBlanksAs.GAP,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(
            org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs.INT_GAP, "gap"));
    assertEquals(
        ExcelChartDisplayBlanksAs.SPAN,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(
            org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs.INT_SPAN, "span"));
    assertEquals(
        ExcelChartDisplayBlanksAs.ZERO,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(
            org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs.INT_ZERO, "zero"));
    IllegalArgumentException unsupportedDisplayBlanks =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelChartPoiBridge.fromPoiDisplayBlanks(99, "bogus"));
    assertTrue(unsupportedDisplayBlanks.getMessage().contains("Unsupported displayBlanksAs token"));
    assertEquals(
        ExcelChartDisplayBlanksAs.GAP,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(STDispBlanksAs.GAP));
    assertEquals(
        ExcelChartDisplayBlanksAs.SPAN,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(STDispBlanksAs.SPAN));
    assertEquals(
        ExcelChartDisplayBlanksAs.ZERO,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(STDispBlanksAs.ZERO));
    assertEquals(
        DisplayBlanks.GAP, ExcelChartPoiBridge.toPoiDisplayBlanks(ExcelChartDisplayBlanksAs.GAP));
    assertEquals(
        DisplayBlanks.SPAN, ExcelChartPoiBridge.toPoiDisplayBlanks(ExcelChartDisplayBlanksAs.SPAN));
    assertEquals(
        DisplayBlanks.ZERO, ExcelChartPoiBridge.toPoiDisplayBlanks(ExcelChartDisplayBlanksAs.ZERO));

    assertEquals(
        ExcelChartAxisPosition.BOTTOM,
        ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.BOTTOM));
    assertEquals(
        ExcelChartAxisPosition.LEFT, ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.LEFT));
    assertEquals(
        ExcelChartAxisPosition.RIGHT, ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.RIGHT));
    assertEquals(
        ExcelChartAxisPosition.TOP, ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.TOP));

    assertEquals(
        ExcelChartAxisCrosses.AUTO_ZERO,
        ExcelChartPoiBridge.fromPoiAxisCrosses(AxisCrosses.AUTO_ZERO));
    assertEquals(
        ExcelChartAxisCrosses.MAX, ExcelChartPoiBridge.fromPoiAxisCrosses(AxisCrosses.MAX));
    assertEquals(
        ExcelChartAxisCrosses.MIN, ExcelChartPoiBridge.fromPoiAxisCrosses(AxisCrosses.MIN));
    assertEquals(
        AxisPosition.BOTTOM, ExcelChartPoiBridge.toPoiAxisPosition(ExcelChartAxisPosition.BOTTOM));
    assertEquals(
        AxisPosition.LEFT, ExcelChartPoiBridge.toPoiAxisPosition(ExcelChartAxisPosition.LEFT));
    assertEquals(
        AxisPosition.RIGHT, ExcelChartPoiBridge.toPoiAxisPosition(ExcelChartAxisPosition.RIGHT));
    assertEquals(
        AxisPosition.TOP, ExcelChartPoiBridge.toPoiAxisPosition(ExcelChartAxisPosition.TOP));
    assertEquals(
        AxisCrosses.AUTO_ZERO,
        ExcelChartPoiBridge.toPoiAxisCrosses(ExcelChartAxisCrosses.AUTO_ZERO));
    assertEquals(AxisCrosses.MAX, ExcelChartPoiBridge.toPoiAxisCrosses(ExcelChartAxisCrosses.MAX));
    assertEquals(AxisCrosses.MIN, ExcelChartPoiBridge.toPoiAxisCrosses(ExcelChartAxisCrosses.MIN));

    assertEquals("AREA", ExcelChartPoiBridge.canonicalPlotTypeToken("XDDFAreaChartData"));
    assertEquals("CUSTOMPLOT", ExcelChartPoiBridge.canonicalPlotTypeToken("CustomPlot"));
    // startsWith("XDDF") true but endsWith("ChartData") false — falls through to toUpperCase.
    assertEquals("XDDFUNKNOWN", ExcelChartPoiBridge.canonicalPlotTypeToken("XDDFUnknown"));
  }

  @Test
  void classifiesConcretePlotFamiliesAndAxisKindsAcrossTheFullModeledSurface() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      int chartIndex = 0;
      for (ExcelChartPlotType plotType : ExcelChartPlotType.values()) {
        XSSFChart chart =
            drawing.createChart(
                drawing.createAnchor(
                    0, 0, 0, 0, 1 + (chartIndex * 6), 1, 6 + (chartIndex * 6), 10));
        chartIndex++;
        XDDFChartData data = createChartData(chart, plotType);
        assertEquals(plotType, ExcelChartPoiBridge.plotType(data));
        assertEquals(plotType.name(), ExcelChartPoiBridge.plotTypeToken(data));
      }

      XSSFChart axisChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 12, 6, 20));
      XDDFCategoryAxis categoryAxis = axisChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = axisChart.createValueAxis(AxisPosition.LEFT);
      XDDFDateAxis dateAxis = axisChart.createDateAxis(AxisPosition.TOP);
      XDDFSeriesAxis seriesAxis = axisChart.createSeriesAxis(AxisPosition.RIGHT);
      assertEquals(ExcelChartAxisKind.CATEGORY, ExcelChartPoiBridge.axisKind(categoryAxis));
      assertEquals(ExcelChartAxisKind.VALUE, ExcelChartPoiBridge.axisKind(valueAxis));
      assertEquals(ExcelChartAxisKind.DATE, ExcelChartPoiBridge.axisKind(dateAxis));
      assertEquals(ExcelChartAxisKind.SERIES, ExcelChartPoiBridge.axisKind(seriesAxis));
    }
  }

  @Test
  void plotTypeMappingMatchesPoiChartTypeEnum() {
    assertEquals(ChartTypes.AREA, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.AREA));
    assertEquals(ChartTypes.AREA3D, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.AREA_3D));
    assertEquals(ChartTypes.BAR, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.BAR));
    assertEquals(ChartTypes.BAR3D, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.BAR_3D));
    assertEquals(
        ChartTypes.DOUGHNUT, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.DOUGHNUT));
    assertEquals(ChartTypes.LINE, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.LINE));
    assertEquals(ChartTypes.LINE3D, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.LINE_3D));
    assertEquals(ChartTypes.PIE, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.PIE));
    assertEquals(ChartTypes.PIE3D, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.PIE_3D));
    assertEquals(ChartTypes.RADAR, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.RADAR));
    assertEquals(
        ChartTypes.SCATTER, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.SCATTER));
    assertEquals(
        ChartTypes.SURFACE, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.SURFACE));
    assertEquals(
        ChartTypes.SURFACE3D, ExcelChartPoiBridge.toPoiChartType(ExcelChartPlotType.SURFACE_3D));
  }

  @Test
  void rejectsUnsupportedAxisAndChartDataFamiliesExplicitly() {
    IllegalArgumentException unsupportedAxis =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelChartPoiBridge.axisKind(new UnsupportedAxis()));
    assertTrue(unsupportedAxis.getMessage().contains("Unsupported chart axis family"));

    UnsupportedChartData unsupportedChartData = new UnsupportedChartData();
    IllegalArgumentException unsupportedPlot =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelChartPoiBridge.plotType(unsupportedChartData));
    assertTrue(unsupportedPlot.getMessage().contains("Unsupported chart data family"));
    assertEquals("UNSUPPORTEDCHARTDATA", ExcelChartPoiBridge.plotTypeToken(unsupportedChartData));
  }

  private static XDDFChartData createChartData(XSSFChart chart, ExcelChartPlotType plotType) {
    return switch (plotType) {
      case AREA -> chart.createData(ChartTypes.AREA, categoryAxis(chart), valueAxis(chart));
      case AREA_3D -> chart.createData(ChartTypes.AREA3D, categoryAxis(chart), valueAxis(chart));
      case BAR -> chart.createData(ChartTypes.BAR, categoryAxis(chart), valueAxis(chart));
      case BAR_3D -> chart.createData(ChartTypes.BAR3D, categoryAxis(chart), valueAxis(chart));
      case DOUGHNUT -> chart.createData(ChartTypes.DOUGHNUT, null, null);
      case LINE -> chart.createData(ChartTypes.LINE, categoryAxis(chart), valueAxis(chart));
      case LINE_3D -> chart.createData(ChartTypes.LINE3D, categoryAxis(chart), valueAxis(chart));
      case PIE -> chart.createData(ChartTypes.PIE, null, null);
      case PIE_3D -> chart.createData(ChartTypes.PIE3D, null, null);
      case RADAR -> chart.createData(ChartTypes.RADAR, categoryAxis(chart), valueAxis(chart));
      case SCATTER ->
          chart.createData(
              ChartTypes.SCATTER,
              valueAxis(chart, AxisPosition.BOTTOM),
              valueAxis(chart, AxisPosition.LEFT));
      case SURFACE -> chart.createData(ChartTypes.SURFACE, categoryAxis(chart), valueAxis(chart));
      case SURFACE_3D ->
          chart.createData(ChartTypes.SURFACE3D, categoryAxis(chart), valueAxis(chart));
    };
  }

  private static XDDFCategoryAxis categoryAxis(XSSFChart chart) {
    return chart.createCategoryAxis(AxisPosition.BOTTOM);
  }

  private static XDDFValueAxis valueAxis(XSSFChart chart) {
    return chart.createValueAxis(AxisPosition.LEFT);
  }

  private static XDDFValueAxis valueAxis(XSSFChart chart, AxisPosition position) {
    return chart.createValueAxis(position);
  }

  /** Minimal unsupported chart-data stub for explicit unsupported-family translation tests. */
  private static final class UnsupportedChartData extends XDDFChartData {
    private UnsupportedChartData() {
      super(null);
    }

    @Override
    protected void removeCTSeries(int index) {}

    @Override
    public void setVaryColors(Boolean varyColors) {}

    @Override
    public Series addSeries(
        XDDFDataSource<?> category, XDDFNumericalDataSource<? extends Number> values) {
      return null;
    }
  }

  /** Minimal unsupported axis stub for explicit unsupported-axis translation tests. */
  private static final class UnsupportedAxis extends XDDFChartAxis {
    private final CTUnsignedInt axisId = CTUnsignedInt.Factory.newInstance();
    private final CTAxPos axisPosition = CTAxPos.Factory.newInstance();
    private final CTNumFmt numberFormat = CTNumFmt.Factory.newInstance();
    private final CTScaling scaling = CTScaling.Factory.newInstance();
    private final CTCrosses crosses = CTCrosses.Factory.newInstance();
    private final CTBoolean delete = CTBoolean.Factory.newInstance();
    private final CTTickMark majorTickMark = CTTickMark.Factory.newInstance();
    private final CTTickMark minorTickMark = CTTickMark.Factory.newInstance();
    private final CTTickLblPos tickLabelPosition = CTTickLblPos.Factory.newInstance();

    @Override
    protected CTUnsignedInt getCTAxId() {
      return axisId;
    }

    @Override
    protected CTAxPos getCTAxPos() {
      return axisPosition;
    }

    @Override
    protected CTNumFmt getCTNumFmt() {
      return numberFormat;
    }

    @Override
    protected CTScaling getCTScaling() {
      return scaling;
    }

    @Override
    protected CTCrosses getCTCrosses() {
      return crosses;
    }

    @Override
    protected CTBoolean getDelete() {
      return delete;
    }

    @Override
    protected CTTickMark getMajorCTTickMark() {
      return majorTickMark;
    }

    @Override
    protected CTTickMark getMinorCTTickMark() {
      return minorTickMark;
    }

    @Override
    protected CTTickLblPos getCTTickLblPos() {
      return tickLabelPosition;
    }

    @Override
    public XDDFShapeProperties getOrAddMajorGridProperties() {
      return null;
    }

    @Override
    public XDDFShapeProperties getOrAddMinorGridProperties() {
      return null;
    }

    @Override
    public XDDFShapeProperties getOrAddShapeProperties() {
      return null;
    }

    @Override
    public XDDFRunProperties getOrAddTextProperties() {
      return null;
    }

    @Override
    public void setTitle(String text) {}

    @Override
    public boolean isSetMinorUnit() {
      return false;
    }

    @Override
    public void setMinorUnit(double minor) {}

    @Override
    public double getMinorUnit() {
      return 0d;
    }

    @Override
    public boolean isSetMajorUnit() {
      return false;
    }

    @Override
    public void setMajorUnit(double major) {}

    @Override
    public double getMajorUnit() {
      return 0d;
    }

    @Override
    public boolean hasNumberFormat() {
      return false;
    }

    @Override
    public void crossAxis(XDDFChartAxis axis) {}
  }

  private static void seedData(XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(18d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
  }
}
