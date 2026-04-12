package dev.erst.gridgrind.excel;

import java.util.Locale;
import java.util.Objects;
import org.apache.poi.xddf.usermodel.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;

/** Package-owned translation seam between POI chart enums/classes and GridGrind chart types. */
final class ExcelChartPoiBridge {
  private ExcelChartPoiBridge() {}

  static ExcelChartAxisKind axisKind(XDDFChartAxis axis) {
    Objects.requireNonNull(axis, "axis must not be null");
    return switch (axis) {
      case XDDFCategoryAxis _ -> ExcelChartAxisKind.CATEGORY;
      case XDDFValueAxis _ -> ExcelChartAxisKind.VALUE;
      default ->
          throw new IllegalArgumentException(
              "Chart axis family is outside the current modeled simple-chart contract: "
                  + axis.getClass().getName());
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

  static ExcelChartAxisCrosses fromPoiAxisCrosses(AxisCrosses crosses) {
    Objects.requireNonNull(crosses, "crosses must not be null");
    return switch (crosses) {
      case AUTO_ZERO -> ExcelChartAxisCrosses.AUTO_ZERO;
      case MAX -> ExcelChartAxisCrosses.MAX;
      case MIN -> ExcelChartAxisCrosses.MIN;
    };
  }

  static String plotTypeToken(XDDFChartData chartData) {
    Objects.requireNonNull(chartData, "chartData must not be null");
    return switch (chartData) {
      case XDDFBarChartData _ -> "BAR";
      case XDDFLineChartData _ -> "LINE";
      case XDDFPieChartData _ -> "PIE";
      default -> canonicalPlotTypeToken(chartData.getClass().getSimpleName());
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
