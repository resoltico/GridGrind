package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

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
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import java.io.IOException;
import java.util.Map;
import java.util.Optional;
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
  void convertsAllModeledPoiChartEnumsAndTokens() {
    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            BarDirection.COL, ExcelChartBarDirection.COLUMN,
            BarDirection.BAR, ExcelChartBarDirection.BAR),
        ChartAdapter::fromPoiBarDirection,
        ChartAdapter::toPoiBarDirection);

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            BarGrouping.STANDARD, ExcelChartBarGrouping.STANDARD,
            BarGrouping.CLUSTERED, ExcelChartBarGrouping.CLUSTERED,
            BarGrouping.STACKED, ExcelChartBarGrouping.STACKED,
            BarGrouping.PERCENT_STACKED, ExcelChartBarGrouping.PERCENT_STACKED),
        ChartAdapter::fromPoiBarGroupingOrDefault,
        ChartAdapter::toPoiBarGrouping);
    assertEquals(ExcelChartBarGrouping.CLUSTERED, ChartAdapter.fromPoiBarGroupingOrDefault(null));
    EnumMappingAssertions.assertMappings(
        Map.of(
            "standard", ExcelChartBarGrouping.STANDARD,
            "clustered", ExcelChartBarGrouping.CLUSTERED,
            "stacked", ExcelChartBarGrouping.STACKED,
            "percentstacked", ExcelChartBarGrouping.PERCENT_STACKED),
        ChartAdapter::fromBarGroupingTokenOrDefault);
    assertEquals(ExcelChartBarGrouping.CLUSTERED, ChartAdapter.fromBarGroupingTokenOrDefault(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ChartAdapter.fromBarGroupingTokenOrDefault("unsupported"));

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            Grouping.STANDARD, ExcelChartGrouping.STANDARD,
            Grouping.STACKED, ExcelChartGrouping.STACKED,
            Grouping.PERCENT_STACKED, ExcelChartGrouping.PERCENT_STACKED),
        ChartAdapter::fromPoiGrouping,
        ChartAdapter::toPoiGrouping);
    assertEquals(ExcelChartGrouping.STANDARD, ChartAdapter.fromPoiGroupingOrDefault(null));
    EnumMappingAssertions.assertMappings(
        Map.of(
            Grouping.STANDARD, ExcelChartGrouping.STANDARD,
            Grouping.STACKED, ExcelChartGrouping.STACKED,
            Grouping.PERCENT_STACKED, ExcelChartGrouping.PERCENT_STACKED),
        ChartAdapter::fromPoiGroupingOrDefault);
    EnumMappingAssertions.assertMappings(
        Map.of(
            "standard", ExcelChartGrouping.STANDARD,
            "stacked", ExcelChartGrouping.STACKED,
            "percentstacked", ExcelChartGrouping.PERCENT_STACKED),
        ChartAdapter::fromGroupingTokenOrDefault);
    assertEquals(ExcelChartGrouping.STANDARD, ChartAdapter.fromGroupingTokenOrDefault(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ChartAdapter.fromGroupingTokenOrDefault("unsupported"));

    EnumMappingAssertions.assertOptionalBidirectionalMappings(
        Map.of(
            Shape.BOX, ExcelChartBarShape.BOX,
            Shape.CONE, ExcelChartBarShape.CONE,
            Shape.CONE_TO_MAX, ExcelChartBarShape.CONE_TO_MAX,
            Shape.CYLINDER, ExcelChartBarShape.CYLINDER,
            Shape.PYRAMID, ExcelChartBarShape.PYRAMID,
            Shape.PYRAMID_TO_MAX, ExcelChartBarShape.PYRAMID_TO_MAX),
        ChartAdapter::fromPoiBarShape,
        ChartAdapter::toPoiBarShape);
    assertEquals(Optional.empty(), ChartAdapter.fromPoiBarShape(null));
    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            RadarStyle.FILLED, ExcelChartRadarStyle.FILLED,
            RadarStyle.MARKER, ExcelChartRadarStyle.MARKER,
            RadarStyle.STANDARD, ExcelChartRadarStyle.STANDARD),
        ChartAdapter::fromPoiRadarStyle,
        ChartAdapter::toPoiRadarStyle);

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            ScatterStyle.LINE, ExcelChartScatterStyle.LINE,
            ScatterStyle.LINE_MARKER, ExcelChartScatterStyle.LINE_MARKER,
            ScatterStyle.MARKER, ExcelChartScatterStyle.MARKER,
            ScatterStyle.NONE, ExcelChartScatterStyle.NONE,
            ScatterStyle.SMOOTH, ExcelChartScatterStyle.SMOOTH,
            ScatterStyle.SMOOTH_MARKER, ExcelChartScatterStyle.SMOOTH_MARKER),
        ChartAdapter::fromPoiScatterStyle,
        ChartAdapter::toPoiScatterStyle);

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.ofEntries(
            Map.entry(MarkerStyle.CIRCLE, ExcelChartMarkerStyle.CIRCLE),
            Map.entry(MarkerStyle.DASH, ExcelChartMarkerStyle.DASH),
            Map.entry(MarkerStyle.DIAMOND, ExcelChartMarkerStyle.DIAMOND),
            Map.entry(MarkerStyle.DOT, ExcelChartMarkerStyle.DOT),
            Map.entry(MarkerStyle.NONE, ExcelChartMarkerStyle.NONE),
            Map.entry(MarkerStyle.PICTURE, ExcelChartMarkerStyle.PICTURE),
            Map.entry(MarkerStyle.PLUS, ExcelChartMarkerStyle.PLUS),
            Map.entry(MarkerStyle.SQUARE, ExcelChartMarkerStyle.SQUARE),
            Map.entry(MarkerStyle.STAR, ExcelChartMarkerStyle.STAR),
            Map.entry(MarkerStyle.TRIANGLE, ExcelChartMarkerStyle.TRIANGLE),
            Map.entry(MarkerStyle.X, ExcelChartMarkerStyle.X)),
        ChartAdapter::fromPoi,
        ChartAdapter::toPoi);

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            LegendPosition.BOTTOM, ExcelChartLegendPosition.BOTTOM,
            LegendPosition.LEFT, ExcelChartLegendPosition.LEFT,
            LegendPosition.RIGHT, ExcelChartLegendPosition.RIGHT,
            LegendPosition.TOP, ExcelChartLegendPosition.TOP,
            LegendPosition.TOP_RIGHT, ExcelChartLegendPosition.TOP_RIGHT),
        ChartAdapter::fromPoiLegendPosition,
        ChartAdapter::toPoiLegendPosition);

    EnumMappingAssertions.assertMappings(
        Map.of(
            STDispBlanksAs.GAP, ExcelChartDisplayBlanksAs.GAP,
            STDispBlanksAs.SPAN, ExcelChartDisplayBlanksAs.SPAN,
            STDispBlanksAs.ZERO, ExcelChartDisplayBlanksAs.ZERO),
        ChartAdapter::fromPoiDisplayBlanks);
    EnumMappingAssertions.assertMappings(
        Map.of(
            ExcelChartDisplayBlanksAs.GAP, DisplayBlanks.GAP,
            ExcelChartDisplayBlanksAs.SPAN, DisplayBlanks.SPAN,
            ExcelChartDisplayBlanksAs.ZERO, DisplayBlanks.ZERO),
        ChartAdapter::toPoiDisplayBlanks);

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            AxisPosition.BOTTOM, ExcelChartAxisPosition.BOTTOM,
            AxisPosition.LEFT, ExcelChartAxisPosition.LEFT,
            AxisPosition.RIGHT, ExcelChartAxisPosition.RIGHT,
            AxisPosition.TOP, ExcelChartAxisPosition.TOP),
        ChartAdapter::fromPoiAxisPosition,
        ChartAdapter::toPoiAxisPosition);

    EnumMappingAssertions.assertBidirectionalMappings(
        Map.of(
            AxisCrosses.AUTO_ZERO, ExcelChartAxisCrosses.AUTO_ZERO,
            AxisCrosses.MAX, ExcelChartAxisCrosses.MAX,
            AxisCrosses.MIN, ExcelChartAxisCrosses.MIN),
        ChartAdapter::fromPoiAxisCrosses,
        ChartAdapter::toPoiAxisCrosses);

    EnumMappingAssertions.assertMappings(
        Map.of(
            "XDDFAreaChartData", "AREA",
            "CustomPlot", "CUSTOMPLOT",
            "XDDFUnknown", "XDDFUNKNOWN"),
        ChartAdapter::canonicalPlotTypeToken);
  }

  @Test
  void classifiesConcretePlotFamiliesAndAxisKindsAcrossTheFullModeledSurface() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      int chartIndex = 0;
      for (var entry :
          Map.ofEntries(
                  Map.entry(ChartTypes.AREA, "AREA"),
                  Map.entry(ChartTypes.AREA3D, "AREA_3D"),
                  Map.entry(ChartTypes.BAR, "BAR"),
                  Map.entry(ChartTypes.BAR3D, "BAR_3D"),
                  Map.entry(ChartTypes.DOUGHNUT, "DOUGHNUT"),
                  Map.entry(ChartTypes.LINE, "LINE"),
                  Map.entry(ChartTypes.LINE3D, "LINE_3D"),
                  Map.entry(ChartTypes.PIE, "PIE"),
                  Map.entry(ChartTypes.PIE3D, "PIE_3D"),
                  Map.entry(ChartTypes.RADAR, "RADAR"),
                  Map.entry(ChartTypes.SCATTER, "SCATTER"),
                  Map.entry(ChartTypes.SURFACE, "SURFACE"),
                  Map.entry(ChartTypes.SURFACE3D, "SURFACE_3D"))
              .entrySet()) {
        ChartTypes plotType = entry.getKey();
        XSSFChart chart =
            drawing.createChart(
                drawing.createAnchor(
                    0, 0, 0, 0, 1 + (chartIndex * 6), 1, 6 + (chartIndex * 6), 10));
        chartIndex++;
        XDDFChartData data = createChartData(chart, plotType);
        assertEquals(entry.getValue(), ChartAdapter.plotTypeToken(data));
      }

      XSSFChart axisChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 12, 6, 20));
      XDDFCategoryAxis categoryAxis = axisChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = axisChart.createValueAxis(AxisPosition.LEFT);
      XDDFDateAxis dateAxis = axisChart.createDateAxis(AxisPosition.TOP);
      XDDFSeriesAxis seriesAxis = axisChart.createSeriesAxis(AxisPosition.RIGHT);
      assertEquals(ExcelChartAxisKind.CATEGORY, ChartAdapter.axisKind(categoryAxis));
      assertEquals(ExcelChartAxisKind.VALUE, ChartAdapter.axisKind(valueAxis));
      assertEquals(ExcelChartAxisKind.DATE, ChartAdapter.axisKind(dateAxis));
      assertEquals(ExcelChartAxisKind.SERIES, ChartAdapter.axisKind(seriesAxis));
    }
  }

  @Test
  void rejectsUnsupportedAxesAndReportsUnknownChartDataTokens() {
    IllegalArgumentException unsupportedAxis =
        assertThrows(
            IllegalArgumentException.class, () -> ChartAdapter.axisKind(new UnsupportedAxis()));
    assertTrue(unsupportedAxis.getMessage().contains("Unsupported chart axis family"));

    UnsupportedChartData unsupportedChartData = new UnsupportedChartData();
    assertEquals("UNSUPPORTEDCHARTDATA", ChartAdapter.plotTypeToken(unsupportedChartData));
  }

  private static XDDFChartData createChartData(XSSFChart chart, ChartTypes plotType) {
    return switch (plotType) {
      case AREA -> chart.createData(ChartTypes.AREA, categoryAxis(chart), valueAxis(chart));
      case AREA3D -> chart.createData(ChartTypes.AREA3D, categoryAxis(chart), valueAxis(chart));
      case BAR -> chart.createData(ChartTypes.BAR, categoryAxis(chart), valueAxis(chart));
      case BAR3D -> chart.createData(ChartTypes.BAR3D, categoryAxis(chart), valueAxis(chart));
      case DOUGHNUT -> chart.createData(ChartTypes.DOUGHNUT, null, null);
      case LINE -> chart.createData(ChartTypes.LINE, categoryAxis(chart), valueAxis(chart));
      case LINE3D -> chart.createData(ChartTypes.LINE3D, categoryAxis(chart), valueAxis(chart));
      case PIE -> chart.createData(ChartTypes.PIE, null, null);
      case PIE3D -> chart.createData(ChartTypes.PIE3D, null, null);
      case RADAR -> chart.createData(ChartTypes.RADAR, categoryAxis(chart), valueAxis(chart));
      case SCATTER ->
          chart.createData(
              ChartTypes.SCATTER,
              valueAxis(chart, AxisPosition.BOTTOM),
              valueAxis(chart, AxisPosition.LEFT));
      case SURFACE -> chart.createData(ChartTypes.SURFACE, categoryAxis(chart), valueAxis(chart));
      case SURFACE3D ->
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
