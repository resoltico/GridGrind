package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
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
import java.util.Optional;

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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartGrouping grouping,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one area plot from the authored wire shape. */
    @JsonCreator
    public Area(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), grouping, axes, series);
    }

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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartGrouping grouping,
      Optional<Integer> gapDepth,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one 3D area plot from the authored wire shape. */
    @JsonCreator
    public Area3D(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("gapDepth") Optional<Integer> gapDepth,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), grouping, emptyIfNull(gapDepth), axes, series);
    }

    public Area3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D area plot with the default category/value axis pair. */
    public Area3D(
        boolean varyColors,
        ExcelChartGrouping grouping,
        Optional<Integer> gapDepth,
        List<ChartSeriesInput> series) {
      this(varyColors, grouping, gapDepth, defaultCategoryAxes(), series);
    }
  }

  /** Bar chart plot. */
  record Bar(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapWidth,
      Optional<Integer> overlap,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one bar plot from the authored wire shape. */
    @JsonCreator
    public Bar(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("barDirection") ExcelChartBarDirection barDirection,
        @JsonProperty("grouping") ExcelChartBarGrouping grouping,
        @JsonProperty("gapWidth") Optional<Integer> gapWidth,
        @JsonProperty("overlap") Optional<Integer> overlap,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(
          defaultFalse(varyColors),
          barDirection,
          grouping,
          emptyIfNull(gapWidth),
          emptyIfNull(overlap),
          axes,
          series);
    }

    public Bar {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapWidth, "gapWidth must not be null");
      Objects.requireNonNull(overlap, "overlap must not be null");
      if (overlap.isPresent() && (overlap.orElseThrow() < -100 || overlap.orElseThrow() > 100)) {
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
        Optional<Integer> gapWidth,
        Optional<Integer> overlap,
        List<ChartSeriesInput> series) {
      this(varyColors, barDirection, grouping, gapWidth, overlap, defaultCategoryAxes(), series);
    }
  }

  /** 3D bar chart plot. */
  record Bar3D(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapDepth,
      Optional<Integer> gapWidth,
      Optional<ExcelChartBarShape> shape,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one 3D bar plot from the authored wire shape. */
    @JsonCreator
    public Bar3D(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("barDirection") ExcelChartBarDirection barDirection,
        @JsonProperty("grouping") ExcelChartBarGrouping grouping,
        @JsonProperty("gapDepth") Optional<Integer> gapDepth,
        @JsonProperty("gapWidth") Optional<Integer> gapWidth,
        @JsonProperty("shape") Optional<ExcelChartBarShape> shape,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(
          defaultFalse(varyColors),
          barDirection,
          grouping,
          emptyIfNull(gapDepth),
          emptyIfNull(gapWidth),
          emptyIfNull(shape),
          axes,
          series);
    }

    public Bar3D {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      Objects.requireNonNull(gapWidth, "gapWidth must not be null");
      Objects.requireNonNull(shape, "shape must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D bar plot with the default category/value axis pair. */
    public Bar3D(
        boolean varyColors,
        ExcelChartBarDirection barDirection,
        ExcelChartBarGrouping grouping,
        Optional<Integer> gapDepth,
        Optional<Integer> gapWidth,
        Optional<ExcelChartBarShape> shape,
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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      Optional<Integer> firstSliceAngle,
      Optional<Integer> holeSize,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one doughnut plot from the authored wire shape. */
    @JsonCreator
    public Doughnut(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("firstSliceAngle") Optional<Integer> firstSliceAngle,
        @JsonProperty("holeSize") Optional<Integer> holeSize,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), emptyIfNull(firstSliceAngle), emptyIfNull(holeSize), series);
    }

    public Doughnut {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      Objects.requireNonNull(holeSize, "holeSize must not be null");
      validateAngle(firstSliceAngle);
      if (holeSize.isPresent() && (holeSize.orElseThrow() < 10 || holeSize.orElseThrow() > 90)) {
        throw new IllegalArgumentException("holeSize must be between 10 and 90");
      }
      series = copySeries(series);
    }
  }

  /** Line chart plot. */
  record Line(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartGrouping grouping,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one line plot from the authored wire shape. */
    @JsonCreator
    public Line(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), grouping, axes, series);
    }

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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartGrouping grouping,
      Optional<Integer> gapDepth,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one 3D line plot from the authored wire shape. */
    @JsonCreator
    public Line3D(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("gapDepth") Optional<Integer> gapDepth,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), grouping, emptyIfNull(gapDepth), axes, series);
    }

    public Line3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D line plot with the default category/value axis pair. */
    public Line3D(
        boolean varyColors,
        ExcelChartGrouping grouping,
        Optional<Integer> gapDepth,
        List<ChartSeriesInput> series) {
      this(varyColors, grouping, gapDepth, defaultCategoryAxes(), series);
    }
  }

  /** Pie chart plot. */
  record Pie(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      Optional<Integer> firstSliceAngle,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one pie plot from the authored wire shape. */
    @JsonCreator
    public Pie(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("firstSliceAngle") Optional<Integer> firstSliceAngle,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), emptyIfNull(firstSliceAngle), series);
    }

    public Pie {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      validateAngle(firstSliceAngle);
      series = copySeries(series);
    }
  }

  /** 3D pie chart plot. */
  record Pie3D(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one 3D pie plot from the authored wire shape. */
    @JsonCreator
    public Pie3D(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), series);
    }

    public Pie3D {
      series = copySeries(series);
    }
  }

  /** Radar chart plot. */
  record Radar(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle style,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one radar plot from the authored wire shape. */
    @JsonCreator
    public Radar(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("style") dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle style,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), style, axes, series);
    }

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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      ExcelChartScatterStyle style,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one scatter plot from the authored wire shape. */
    @JsonCreator
    public Scatter(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("style") ExcelChartScatterStyle style,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), style, axes, series);
    }

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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean wireframe,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one surface plot from the authored wire shape. */
    @JsonCreator
    public Surface(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("wireframe") Boolean wireframe,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), defaultFalse(wireframe), axes, series);
    }

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
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean varyColors,
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean wireframe,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {
    /** Reads one 3D surface plot from the authored wire shape. */
    @JsonCreator
    public Surface3D(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("wireframe") Boolean wireframe,
        @JsonProperty("axes") List<ChartAxisInput> axes,
        @JsonProperty("series") List<ChartSeriesInput> series) {
      this(defaultFalse(varyColors), defaultFalse(wireframe), axes, series);
    }

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

  private static boolean defaultFalse(Boolean value) {
    return ProtocolBooleanDefault.FALSE.resolve(value);
  }

  private static <T> Optional<T> emptyIfNull(Optional<T> value) {
    return value == null ? Optional.empty() : value;
  }

  private static void validateAngle(Optional<Integer> angle) {
    if (angle.isPresent() && (angle.orElseThrow() < 0 || angle.orElseThrow() > 360)) {
      throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
    }
  }
}
