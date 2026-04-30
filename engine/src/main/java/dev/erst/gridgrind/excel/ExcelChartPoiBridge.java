package dev.erst.gridgrind.excel;

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
import java.util.Locale;
import java.util.Objects;
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
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;

/**
 * Package-owned translation seam between POI chart enums/classes and GridGrind chart types.
 *
 * <p>The import count is the direct shape of the supported POI chart enum surface.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class ExcelChartPoiBridge {
  private ExcelChartPoiBridge() {}

  static ExcelChartAxisKind axisKind(XDDFChartAxis axis) {
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

  static XDDFChartAxis createAxis(
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

  static ChartTypes toPoiChartType(ExcelChartPlotType plotType) {
    Objects.requireNonNull(plotType, "plotType must not be null");
    return switch (plotType) {
      case AREA -> ChartTypes.AREA;
      case AREA_3D -> ChartTypes.AREA3D;
      case BAR -> ChartTypes.BAR;
      case BAR_3D -> ChartTypes.BAR3D;
      case DOUGHNUT -> ChartTypes.DOUGHNUT;
      case LINE -> ChartTypes.LINE;
      case LINE_3D -> ChartTypes.LINE3D;
      case PIE -> ChartTypes.PIE;
      case PIE_3D -> ChartTypes.PIE3D;
      case RADAR -> ChartTypes.RADAR;
      case SCATTER -> ChartTypes.SCATTER;
      case SURFACE -> ChartTypes.SURFACE;
      case SURFACE_3D -> ChartTypes.SURFACE3D;
    };
  }

  static ExcelChartPlotType plotType(XDDFChartData chartData) {
    Objects.requireNonNull(chartData, "chartData must not be null");
    return switch (chartData) {
      case XDDFAreaChartData _ -> ExcelChartPlotType.AREA;
      case XDDFArea3DChartData _ -> ExcelChartPlotType.AREA_3D;
      case XDDFBarChartData _ -> ExcelChartPlotType.BAR;
      case XDDFBar3DChartData _ -> ExcelChartPlotType.BAR_3D;
      case XDDFDoughnutChartData _ -> ExcelChartPlotType.DOUGHNUT;
      case XDDFLineChartData _ -> ExcelChartPlotType.LINE;
      case XDDFLine3DChartData _ -> ExcelChartPlotType.LINE_3D;
      case XDDFPieChartData _ -> ExcelChartPlotType.PIE;
      case XDDFPie3DChartData _ -> ExcelChartPlotType.PIE_3D;
      case XDDFRadarChartData _ -> ExcelChartPlotType.RADAR;
      case XDDFScatterChartData _ -> ExcelChartPlotType.SCATTER;
      case XDDFSurfaceChartData _ -> ExcelChartPlotType.SURFACE;
      case XDDFSurface3DChartData _ -> ExcelChartPlotType.SURFACE_3D;
      default -> throw new IllegalArgumentException("Unsupported chart data family: " + chartData);
    };
  }

  static String plotTypeToken(XDDFChartData chartData) {
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

  static ExcelChartBarDirection fromPoiBarDirection(BarDirection direction) {
    Objects.requireNonNull(direction, "direction must not be null");
    return switch (direction) {
      case COL -> ExcelChartBarDirection.COLUMN;
      case BAR -> ExcelChartBarDirection.BAR;
    };
  }

  static BarDirection toPoiBarDirection(ExcelChartBarDirection direction) {
    Objects.requireNonNull(direction, "direction must not be null");
    return switch (direction) {
      case COLUMN -> BarDirection.COL;
      case BAR -> BarDirection.BAR;
    };
  }

  static ExcelChartBarGrouping fromPoiBarGrouping(BarGrouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> ExcelChartBarGrouping.STANDARD;
      case CLUSTERED -> ExcelChartBarGrouping.CLUSTERED;
      case STACKED -> ExcelChartBarGrouping.STACKED;
      case PERCENT_STACKED -> ExcelChartBarGrouping.PERCENT_STACKED;
    };
  }

  static ExcelChartBarGrouping fromPoiBarGroupingOrDefault(BarGrouping grouping) {
    return grouping == null ? ExcelChartBarGrouping.CLUSTERED : fromPoiBarGrouping(grouping);
  }

  static ExcelChartBarGrouping fromBarGroupingTokenOrDefault(String token) {
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

  static BarGrouping toPoiBarGrouping(ExcelChartBarGrouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> BarGrouping.STANDARD;
      case CLUSTERED -> BarGrouping.CLUSTERED;
      case STACKED -> BarGrouping.STACKED;
      case PERCENT_STACKED -> BarGrouping.PERCENT_STACKED;
    };
  }

  static ExcelChartGrouping fromPoiGrouping(Grouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> ExcelChartGrouping.STANDARD;
      case STACKED -> ExcelChartGrouping.STACKED;
      case PERCENT_STACKED -> ExcelChartGrouping.PERCENT_STACKED;
    };
  }

  static ExcelChartGrouping fromPoiGroupingOrDefault(Grouping grouping) {
    return grouping == null ? ExcelChartGrouping.STANDARD : fromPoiGrouping(grouping);
  }

  static ExcelChartGrouping fromGroupingTokenOrDefault(String token) {
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

  static Grouping toPoiGrouping(ExcelChartGrouping grouping) {
    Objects.requireNonNull(grouping, "grouping must not be null");
    return switch (grouping) {
      case STANDARD -> Grouping.STANDARD;
      case STACKED -> Grouping.STACKED;
      case PERCENT_STACKED -> Grouping.PERCENT_STACKED;
    };
  }

  static ExcelChartBarShape fromPoiBarShape(Shape shape) {
    Objects.requireNonNull(shape, "shape must not be null");
    return switch (shape) {
      case BOX -> ExcelChartBarShape.BOX;
      case CONE -> ExcelChartBarShape.CONE;
      case CONE_TO_MAX -> ExcelChartBarShape.CONE_TO_MAX;
      case CYLINDER -> ExcelChartBarShape.CYLINDER;
      case PYRAMID -> ExcelChartBarShape.PYRAMID;
      case PYRAMID_TO_MAX -> ExcelChartBarShape.PYRAMID_TO_MAX;
    };
  }

  static Shape toPoiBarShape(ExcelChartBarShape shape) {
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

  static ExcelChartRadarStyle fromPoiRadarStyle(RadarStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case FILLED -> ExcelChartRadarStyle.FILLED;
      case MARKER -> ExcelChartRadarStyle.MARKER;
      case STANDARD -> ExcelChartRadarStyle.STANDARD;
    };
  }

  static RadarStyle toPoiRadarStyle(ExcelChartRadarStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case FILLED -> RadarStyle.FILLED;
      case MARKER -> RadarStyle.MARKER;
      case STANDARD -> RadarStyle.STANDARD;
    };
  }

  static ExcelChartScatterStyle fromPoiScatterStyle(ScatterStyle style) {
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

  static ScatterStyle toPoiScatterStyle(ExcelChartScatterStyle style) {
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

  static ExcelChartMarkerStyle fromPoiMarkerStyle(MarkerStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case CIRCLE -> ExcelChartMarkerStyle.CIRCLE;
      case DASH -> ExcelChartMarkerStyle.DASH;
      case DIAMOND -> ExcelChartMarkerStyle.DIAMOND;
      case DOT -> ExcelChartMarkerStyle.DOT;
      case NONE -> ExcelChartMarkerStyle.NONE;
      case PICTURE -> ExcelChartMarkerStyle.PICTURE;
      case PLUS -> ExcelChartMarkerStyle.PLUS;
      case SQUARE -> ExcelChartMarkerStyle.SQUARE;
      case STAR -> ExcelChartMarkerStyle.STAR;
      case TRIANGLE -> ExcelChartMarkerStyle.TRIANGLE;
      case X -> ExcelChartMarkerStyle.X;
    };
  }

  static MarkerStyle toPoiMarkerStyle(ExcelChartMarkerStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    return switch (style) {
      case CIRCLE -> MarkerStyle.CIRCLE;
      case DASH -> MarkerStyle.DASH;
      case DIAMOND -> MarkerStyle.DIAMOND;
      case DOT -> MarkerStyle.DOT;
      case NONE -> MarkerStyle.NONE;
      case PICTURE -> MarkerStyle.PICTURE;
      case PLUS -> MarkerStyle.PLUS;
      case SQUARE -> MarkerStyle.SQUARE;
      case STAR -> MarkerStyle.STAR;
      case TRIANGLE -> MarkerStyle.TRIANGLE;
      case X -> MarkerStyle.X;
    };
  }

  static ExcelChartLegendPosition fromPoiLegendPosition(LegendPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> ExcelChartLegendPosition.BOTTOM;
      case LEFT -> ExcelChartLegendPosition.LEFT;
      case RIGHT -> ExcelChartLegendPosition.RIGHT;
      case TOP -> ExcelChartLegendPosition.TOP;
      case TOP_RIGHT -> ExcelChartLegendPosition.TOP_RIGHT;
    };
  }

  static LegendPosition toPoiLegendPosition(ExcelChartLegendPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> LegendPosition.BOTTOM;
      case LEFT -> LegendPosition.LEFT;
      case RIGHT -> LegendPosition.RIGHT;
      case TOP -> LegendPosition.TOP;
      case TOP_RIGHT -> LegendPosition.TOP_RIGHT;
    };
  }

  static ExcelChartDisplayBlanksAs fromPoiDisplayBlanks(STDispBlanksAs.Enum displayBlanks) {
    Objects.requireNonNull(displayBlanks, "displayBlanks must not be null");
    return fromPoiDisplayBlanks(displayBlanks.intValue(), displayBlanks.toString());
  }

  static ExcelChartDisplayBlanksAs fromPoiDisplayBlanks(int displayBlanksToken, String rawToken) {
    return switch (displayBlanksToken) {
      case STDispBlanksAs.INT_GAP -> ExcelChartDisplayBlanksAs.GAP;
      case STDispBlanksAs.INT_SPAN -> ExcelChartDisplayBlanksAs.SPAN;
      case STDispBlanksAs.INT_ZERO -> ExcelChartDisplayBlanksAs.ZERO;
      default ->
          throw new IllegalArgumentException("Unsupported displayBlanksAs token: " + rawToken);
    };
  }

  static DisplayBlanks toPoiDisplayBlanks(ExcelChartDisplayBlanksAs displayBlanksAs) {
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    return switch (displayBlanksAs) {
      case GAP -> DisplayBlanks.GAP;
      case SPAN -> DisplayBlanks.SPAN;
      case ZERO -> DisplayBlanks.ZERO;
    };
  }

  static ExcelChartAxisPosition fromPoiAxisPosition(AxisPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> ExcelChartAxisPosition.BOTTOM;
      case LEFT -> ExcelChartAxisPosition.LEFT;
      case RIGHT -> ExcelChartAxisPosition.RIGHT;
      case TOP -> ExcelChartAxisPosition.TOP;
    };
  }

  static AxisPosition toPoiAxisPosition(ExcelChartAxisPosition position) {
    Objects.requireNonNull(position, "position must not be null");
    return switch (position) {
      case BOTTOM -> AxisPosition.BOTTOM;
      case LEFT -> AxisPosition.LEFT;
      case RIGHT -> AxisPosition.RIGHT;
      case TOP -> AxisPosition.TOP;
    };
  }

  static ExcelChartAxisCrosses fromPoiAxisCrosses(AxisCrosses crosses) {
    Objects.requireNonNull(crosses, "crosses must not be null");
    return switch (crosses) {
      case AUTO_ZERO -> ExcelChartAxisCrosses.AUTO_ZERO;
      case MAX -> ExcelChartAxisCrosses.MAX;
      case MIN -> ExcelChartAxisCrosses.MIN;
    };
  }

  static AxisCrosses toPoiAxisCrosses(ExcelChartAxisCrosses crosses) {
    Objects.requireNonNull(crosses, "crosses must not be null");
    return switch (crosses) {
      case AUTO_ZERO -> AxisCrosses.AUTO_ZERO;
      case MAX -> AxisCrosses.MAX;
      case MIN -> AxisCrosses.MIN;
    };
  }

  static String canonicalPlotTypeToken(String simpleName) {
    Objects.requireNonNull(simpleName, "simpleName must not be null");
    if (simpleName.startsWith("XDDF") && simpleName.endsWith("ChartData")) {
      return simpleName
          .substring(4, simpleName.length() - "ChartData".length())
          .toUpperCase(Locale.ROOT);
    }
    return simpleName.toUpperCase(Locale.ROOT);
  }
}
