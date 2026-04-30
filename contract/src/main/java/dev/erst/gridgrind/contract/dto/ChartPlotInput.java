package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import java.util.List;
import java.util.Objects;

/** One authored chart plot. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartPlotInput.Area.class, name = "AREA"),
  @JsonSubTypes.Type(value = ChartPlotInput.Area3D.class, name = "AREA_3D"),
  @JsonSubTypes.Type(value = ChartPlotInput.Bar.class, name = "BAR"),
  @JsonSubTypes.Type(value = ChartPlotInput.Bar3D.class, name = "BAR_3D"),
  @JsonSubTypes.Type(value = ChartPlotInput.Doughnut.class, name = "DOUGHNUT"),
  @JsonSubTypes.Type(value = ChartPlotInput.Line.class, name = "LINE"),
  @JsonSubTypes.Type(value = ChartPlotInput.Line3D.class, name = "LINE_3D"),
  @JsonSubTypes.Type(value = ChartPlotInput.Pie.class, name = "PIE"),
  @JsonSubTypes.Type(value = ChartPlotInput.Pie3D.class, name = "PIE_3D"),
  @JsonSubTypes.Type(value = ChartPlotInput.Radar.class, name = "RADAR"),
  @JsonSubTypes.Type(value = ChartPlotInput.Scatter.class, name = "SCATTER"),
  @JsonSubTypes.Type(value = ChartPlotInput.Surface.class, name = "SURFACE"),
  @JsonSubTypes.Type(value = ChartPlotInput.Surface3D.class, name = "SURFACE_3D")
})
public sealed interface ChartPlotInput
    permits ChartPlotInput.Area,
        ChartPlotInput.Area3D,
        ChartPlotInput.Bar,
        ChartPlotInput.Bar3D,
        ChartPlotInput.Doughnut,
        ChartPlotInput.Line,
        ChartPlotInput.Line3D,
        ChartPlotInput.Pie,
        ChartPlotInput.Pie3D,
        ChartPlotInput.Radar,
        ChartPlotInput.Scatter,
        ChartPlotInput.Surface,
        ChartPlotInput.Surface3D {
  /** Area chart plot. */
  record Area(
      boolean varyColors,
      ExcelChartGrouping grouping,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Area {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one area plot with the default category/value axis pair. */
    public Area(boolean varyColors, ExcelChartGrouping grouping, List<ChartSeriesInput> series) {
      this(varyColors, grouping, defaultCategoryAxes(), series);
    }
  }

  /** 3D area chart plot. */
  record Area3D(
      boolean varyColors,
      ExcelChartGrouping grouping,
      Integer gapDepth,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Area3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D area plot with the default category/value axis pair. */
    public Area3D(
        boolean varyColors,
        ExcelChartGrouping grouping,
        Integer gapDepth,
        List<ChartSeriesInput> series) {
      this(varyColors, grouping, gapDepth, defaultCategoryAxes(), series);
    }
  }

  /** Bar chart plot. */
  record Bar(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Integer gapWidth,
      Integer overlap,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Bar {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      if (overlap != null && (overlap < -100 || overlap > 100)) {
        throw new IllegalArgumentException("overlap must be between -100 and 100");
      }
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one bar plot with the default category/value axis pair. */
    public Bar(
        boolean varyColors,
        ExcelChartBarDirection barDirection,
        ExcelChartBarGrouping grouping,
        Integer gapWidth,
        Integer overlap,
        List<ChartSeriesInput> series) {
      this(varyColors, barDirection, grouping, gapWidth, overlap, defaultCategoryAxes(), series);
    }
  }

  /** 3D bar chart plot. */
  record Bar3D(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Integer gapDepth,
      Integer gapWidth,
      ExcelChartBarShape shape,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Bar3D {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D bar plot with the default category/value axis pair. */
    public Bar3D(
        boolean varyColors,
        ExcelChartBarDirection barDirection,
        ExcelChartBarGrouping grouping,
        Integer gapDepth,
        Integer gapWidth,
        ExcelChartBarShape shape,
        List<ChartSeriesInput> series) {
      this(
          varyColors,
          barDirection,
          grouping,
          gapDepth,
          gapWidth,
          shape,
          defaultCategoryAxes(),
          series);
    }
  }

  /** Doughnut chart plot. */
  record Doughnut(
      boolean varyColors, Integer firstSliceAngle, Integer holeSize, List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Doughnut {
      validateAngle(firstSliceAngle);
      if (holeSize != null && (holeSize < 10 || holeSize > 90)) {
        throw new IllegalArgumentException("holeSize must be between 10 and 90");
      }
      series = copySeries(series);
    }
  }

  /** Line chart plot. */
  record Line(
      boolean varyColors,
      ExcelChartGrouping grouping,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Line {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one line plot with the default category/value axis pair. */
    public Line(boolean varyColors, ExcelChartGrouping grouping, List<ChartSeriesInput> series) {
      this(varyColors, grouping, defaultCategoryAxes(), series);
    }
  }

  /** 3D line chart plot. */
  record Line3D(
      boolean varyColors,
      ExcelChartGrouping grouping,
      Integer gapDepth,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Line3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D line plot with the default category/value axis pair. */
    public Line3D(
        boolean varyColors,
        ExcelChartGrouping grouping,
        Integer gapDepth,
        List<ChartSeriesInput> series) {
      this(varyColors, grouping, gapDepth, defaultCategoryAxes(), series);
    }
  }

  /** Pie chart plot. */
  record Pie(boolean varyColors, Integer firstSliceAngle, List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Pie {
      validateAngle(firstSliceAngle);
      series = copySeries(series);
    }
  }

  /** 3D pie chart plot. */
  record Pie3D(boolean varyColors, List<ChartSeriesInput> series) implements ChartPlotInput {
    public Pie3D {
      series = copySeries(series);
    }
  }

  /** Radar chart plot. */
  record Radar(
      boolean varyColors,
      dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle style,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Radar {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one radar plot with the default category/value axis pair. */
    public Radar(boolean varyColors, ExcelChartRadarStyle style, List<ChartSeriesInput> series) {
      this(varyColors, style, defaultCategoryAxes(), series);
    }
  }

  /** Scatter chart plot. */
  record Scatter(
      boolean varyColors,
      ExcelChartScatterStyle style,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Scatter {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one scatter plot with the default X/Y axis pair. */
    public Scatter(
        boolean varyColors, ExcelChartScatterStyle style, List<ChartSeriesInput> series) {
      this(varyColors, style, defaultScatterAxes(), series);
    }
  }

  /** Surface chart plot. */
  record Surface(
      boolean varyColors,
      boolean wireframe,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Surface {
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one surface plot with the default category/value/series axes. */
    public Surface(boolean varyColors, boolean wireframe, List<ChartSeriesInput> series) {
      this(varyColors, wireframe, defaultSurfaceAxes(), series);
    }
  }

  /** 3D surface chart plot. */
  record Surface3D(
      boolean varyColors,
      boolean wireframe,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    public Surface3D {
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D surface plot with the default category/value/series axes. */
    public Surface3D(boolean varyColors, boolean wireframe, List<ChartSeriesInput> series) {
      this(varyColors, wireframe, defaultSurfaceAxes(), series);
    }
  }

  private static List<ChartAxisInput> defaultCategoryAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisInput> defaultScatterAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisInput> defaultSurfaceAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartSeriesInput> copySeries(List<ChartSeriesInput> series) {
    return ChartInput.copyNonEmptyValues(series, "series");
  }

  private static List<ChartAxisInput> copyAxes(List<ChartAxisInput> axes, String fieldName) {
    return ChartInput.copyNonEmptyValues(axes, fieldName);
  }

  private static void validateAngle(Integer angle) {
    if (angle != null && (angle < 0 || angle > 360)) {
      throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
    }
  }
}
