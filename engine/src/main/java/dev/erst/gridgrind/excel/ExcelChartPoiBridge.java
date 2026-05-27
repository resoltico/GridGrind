package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.DisplayBlanks;
import org.apache.poi.xddf.usermodel.chart.Grouping;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.RadarStyle;
import org.apache.poi.xddf.usermodel.chart.ScatterStyle;
import org.apache.poi.xddf.usermodel.chart.Shape;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;

/**
 * Package-owned translation seam between POI chart enums/classes and GridGrind chart types.
 *
 * <p>The import count is the direct shape of the supported POI chart enum surface.
 */
final class ExcelChartPoiBridge {
  private ExcelChartPoiBridge() {}

  static ExcelChartBarDirection fromPoiBarDirection(BarDirection direction) {
    Objects.requireNonNull(direction, "direction must not be null");
    if (direction == BarDirection.COL) {
      return ExcelChartBarDirection.COLUMN;
    }
    return ExcelChartBarDirection.BAR;
  }

  static BarDirection toPoiBarDirection(ExcelChartBarDirection direction) {
    Objects.requireNonNull(direction, "direction must not be null");
    if (direction == ExcelChartBarDirection.COLUMN) {
      return BarDirection.COL;
    }
    return BarDirection.BAR;
  }

  static ExcelChartBarGrouping fromPoiBarGroupingOrDefault(BarGrouping grouping) {
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

  static ExcelChartBarGrouping fromBarGroupingTokenOrDefault(@Nullable String token) {
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

  static ExcelChartGrouping fromGroupingTokenOrDefault(@Nullable String token) {
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

  static Optional<ExcelChartBarShape> fromPoiBarShape(@Nullable Shape shape) {
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
    if (displayBlanks == STDispBlanksAs.GAP) {
      return ExcelChartDisplayBlanksAs.GAP;
    }
    if (displayBlanks == STDispBlanksAs.SPAN) {
      return ExcelChartDisplayBlanksAs.SPAN;
    }
    return ExcelChartDisplayBlanksAs.ZERO;
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
}
