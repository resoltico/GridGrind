package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** One authored chart plot. */
@Union(discriminator = "type", catalog = true, displayOrder = 49)
public sealed interface ChartPlotInput
    permits dev.erst.gridgrind.core.model.chart.ChartPlotInput.Area,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Area3D,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Bar,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Bar3D,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Doughnut,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Line,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Line3D,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Pie,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Pie3D,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Radar,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Scatter,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Surface,
        dev.erst.gridgrind.core.model.chart.ChartPlotInput.Surface3D {
  /** Area chart plot. */
  @Variant(
      id = "AREA",
      summary =
          "Authored area-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.")
  record Area(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
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
  @Variant(
      id = "AREA_3D",
      summary =
          "Authored 3D area-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.")
  record Area3D(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      ExcelChartGrouping grouping,
      Optional<Integer> gapDepth,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

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
  @Variant(
      id = "BAR",
      summary =
          "Authored bar-chart plot. Supply explicit barDirection, grouping, and axes. The Java authoring surface exposes one convenience constructor for the standard category/value axis pair.")
  record Bar(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapWidth,
      Optional<Integer> overlap,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

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
  @Variant(
      id = "BAR_3D",
      summary =
          "Authored 3D bar-chart plot. Supply explicit barDirection, grouping, and axes. The Java authoring surface exposes one convenience constructor for the standard category/value axis pair.")
  record Bar3D(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapDepth,
      Optional<Integer> gapWidth,
      Optional<ExcelChartBarShape> shape,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

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
  @Variant(
      id = "DOUGHNUT",
      summary =
          "Authored doughnut-chart plot. varyColors, firstSliceAngle, and holeSize are optional.")
  record Doughnut(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      Optional<Integer> firstSliceAngle,
      Optional<Integer> holeSize,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

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
  @Variant(
      id = "LINE",
      summary =
          "Authored line-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.")
  record Line(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
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
  @Variant(
      id = "LINE_3D",
      summary =
          "Authored 3D line-chart plot. Supply explicit grouping and axes, or use the Java convenience constructor for the standard category/value axis pair.")
  record Line3D(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      ExcelChartGrouping grouping,
      Optional<Integer> gapDepth,
      List<ChartAxisInput> axes,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

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
  @Variant(
      id = "PIE",
      summary = "Authored pie-chart plot. varyColors and firstSliceAngle are optional.")
  record Pie(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      Optional<Integer> firstSliceAngle,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

    public Pie {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      validateAngle(firstSliceAngle);
      series = copySeries(series);
    }
  }

  /** 3D pie chart plot. */
  @Variant(id = "PIE_3D", summary = "Authored 3D pie-chart plot.")
  record Pie3D(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      List<ChartSeriesInput> series)
      implements ChartPlotInput {

    public Pie3D {
      series = copySeries(series);
    }
  }

  /** Radar chart plot. */
  @Variant(
      id = "RADAR",
      summary =
          "Authored radar-chart plot. Supply explicit style and axes, or use the Java convenience constructor for the standard category/value axis pair.")
  record Radar(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle style,
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
  @Variant(
      id = "SCATTER",
      summary =
          "Authored scatter-chart plot. Supply explicit style and axes, or use the Java convenience constructor for the standard X/Y axis pair.")
  record Scatter(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
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
  @Variant(
      id = "SURFACE",
      summary =
          "Authored surface-chart plot. Supply explicit axes, or use the Java convenience constructor for the standard category/value/series axis set.")
  record Surface(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      @ProtocolField(optional = true) @WireDefault("false") boolean wireframe,
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
  @Variant(
      id = "SURFACE_3D",
      summary =
          "Authored 3D surface-chart plot. Supply explicit axes, or use the Java convenience constructor for the standard category/value/series axis set.")
  record Surface3D(
      @ProtocolField(optional = true) @WireDefault("false") boolean varyColors,
      @ProtocolField(optional = true) @WireDefault("false") boolean wireframe,
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

  private static void validateAngle(Optional<Integer> angle) {
    if (angle.isPresent() && (angle.orElseThrow() < 0 || angle.orElseThrow() > 360)) {
      throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
    }
  }
}
