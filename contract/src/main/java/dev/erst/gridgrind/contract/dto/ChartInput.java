package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
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
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Authored chart input. One chart may contain one or more plots. */
public record ChartInput(
    String name,
    DrawingAnchorInput anchor,
    Title title,
    Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    Boolean plotOnlyVisibleCells,
    List<Plot> plots) {
  @JsonCreator
  static ChartInput create(
      @JsonProperty("name") String name,
      @JsonProperty("anchor") DrawingAnchorInput anchor,
      @JsonProperty("title") Title title,
      @JsonProperty("legend") Legend legend,
      @JsonProperty("displayBlanksAs") ExcelChartDisplayBlanksAs displayBlanksAs,
      @JsonProperty("plotOnlyVisibleCells") Boolean plotOnlyVisibleCells,
      @JsonProperty("plots") List<Plot> plots) {
    return new ChartInput(
        name,
        anchor,
        title == null ? new Title.None() : title,
        legend == null ? new Legend.Visible(ExcelChartLegendPosition.RIGHT) : legend,
        displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs,
        plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells,
        plots);
  }

  /** Creates one chart with the standard visible legend and gap display defaults. */
  public ChartInput(String name, DrawingAnchorInput anchor, List<Plot> plots) {
    this(
        name,
        anchor,
        new Title.None(),
        new Legend.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.GAP,
        Boolean.TRUE,
        plots);
  }

  public ChartInput {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    Objects.requireNonNull(plotOnlyVisibleCells, "plotOnlyVisibleCells must not be null");
    plots = copyNonEmptyValues(plots, "plots");
  }

  /** Requested chart title definition. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Title.None.class, name = "NONE"),
    @JsonSubTypes.Type(value = Title.Text.class, name = "TEXT"),
    @JsonSubTypes.Type(value = Title.Formula.class, name = "FORMULA")
  })
  public sealed interface Title permits Title.None, Title.Text, Title.Formula {

    /** Remove any chart or series title. */
    record None() implements Title {}

    /** Use a static text title. */
    record Text(TextSourceInput source) implements Title {
      public Text {
        Objects.requireNonNull(source, "source must not be null");
        if (source instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
          throw new IllegalArgumentException("source must not be blank");
        }
      }
    }

    /** Bind the title to one workbook formula. */
    record Formula(String formula) implements Title {
      public Formula {
        formula = requireNonBlank(formula, "formula");
      }
    }
  }

  /** Requested chart-legend state. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Legend.Hidden.class, name = "HIDDEN"),
    @JsonSubTypes.Type(value = Legend.Visible.class, name = "VISIBLE")
  })
  public sealed interface Legend permits Legend.Hidden, Legend.Visible {

    /** Hide the legend entirely. */
    record Hidden() implements Legend {}

    /** Show the legend at one position. */
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** Authored axis definition used by one plot. */
  public record Axis(
      ExcelChartAxisKind kind,
      ExcelChartAxisPosition position,
      ExcelChartAxisCrosses crosses,
      Boolean visible) {
    @JsonCreator
    static Axis create(
        @JsonProperty("kind") ExcelChartAxisKind kind,
        @JsonProperty("position") ExcelChartAxisPosition position,
        @JsonProperty("crosses") ExcelChartAxisCrosses crosses,
        @JsonProperty("visible") Boolean visible) {
      return new Axis(kind, position, crosses, visible == null ? Boolean.TRUE : visible);
    }

    public Axis {
      Objects.requireNonNull(kind, "kind must not be null");
      Objects.requireNonNull(position, "position must not be null");
      Objects.requireNonNull(crosses, "crosses must not be null");
      Objects.requireNonNull(visible, "visible must not be null");
    }

    /** Creates one visible axis without requiring callers to repeat the visibility default. */
    public Axis(
        ExcelChartAxisKind kind, ExcelChartAxisPosition position, ExcelChartAxisCrosses crosses) {
      this(kind, position, crosses, Boolean.TRUE);
    }
  }

  /** One authored chart series. */
  public record Series(
      Title title,
      DataSource categories,
      DataSource values,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      Long explosion) {
    @JsonCreator
    static Series create(
        @JsonProperty("title") Title title,
        @JsonProperty("categories") DataSource categories,
        @JsonProperty("values") DataSource values,
        @JsonProperty("smooth") Boolean smooth,
        @JsonProperty("markerStyle") ExcelChartMarkerStyle markerStyle,
        @JsonProperty("markerSize") Short markerSize,
        @JsonProperty("explosion") Long explosion) {
      return new Series(
          title == null ? new Title.None() : title,
          categories,
          values,
          smooth,
          markerStyle,
          markerSize,
          explosion);
    }

    public Series {
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(categories, "categories must not be null");
      Objects.requireNonNull(values, "values must not be null");
      if (markerSize != null && (markerSize < 2 || markerSize > 72)) {
        throw new IllegalArgumentException("markerSize must be between 2 and 72");
      }
      if (explosion != null && explosion < 0L) {
        throw new IllegalArgumentException("explosion must not be negative");
      }
    }

    /** Creates one series with no explicit title. */
    public Series(
        DataSource categories,
        DataSource values,
        Boolean smooth,
        ExcelChartMarkerStyle markerStyle,
        Short markerSize,
        Long explosion) {
      this(new Title.None(), categories, values, smooth, markerStyle, markerSize, explosion);
    }
  }

  /** Authored chart data source. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = DataSource.Reference.class, name = "REFERENCE"),
    @JsonSubTypes.Type(value = DataSource.StringLiteral.class, name = "STRING_LITERAL"),
    @JsonSubTypes.Type(value = DataSource.NumericLiteral.class, name = "NUMERIC_LITERAL")
  })
  public sealed interface DataSource
      permits DataSource.Reference, DataSource.StringLiteral, DataSource.NumericLiteral {

    /** Workbook-backed chart source formula or defined name. */
    record Reference(String formula) implements DataSource {
      public Reference {
        formula = requireNonBlank(formula, "formula");
      }
    }

    /** Literal string values stored directly in the chart part. */
    record StringLiteral(List<String> values) implements DataSource {
      public StringLiteral {
        values = copyValues(values, "values");
      }
    }

    /** Literal numeric values stored directly in the chart part. */
    record NumericLiteral(List<Double> values) implements DataSource {
      public NumericLiteral {
        values = copyValues(values, "values");
      }
    }
  }

  /** One authored chart plot. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Area.class, name = "AREA"),
    @JsonSubTypes.Type(value = Area3D.class, name = "AREA_3D"),
    @JsonSubTypes.Type(value = Bar.class, name = "BAR"),
    @JsonSubTypes.Type(value = Bar3D.class, name = "BAR_3D"),
    @JsonSubTypes.Type(value = Doughnut.class, name = "DOUGHNUT"),
    @JsonSubTypes.Type(value = Line.class, name = "LINE"),
    @JsonSubTypes.Type(value = Line3D.class, name = "LINE_3D"),
    @JsonSubTypes.Type(value = Pie.class, name = "PIE"),
    @JsonSubTypes.Type(value = Pie3D.class, name = "PIE_3D"),
    @JsonSubTypes.Type(value = Radar.class, name = "RADAR"),
    @JsonSubTypes.Type(value = Scatter.class, name = "SCATTER"),
    @JsonSubTypes.Type(value = Surface.class, name = "SURFACE"),
    @JsonSubTypes.Type(value = Surface3D.class, name = "SURFACE_3D")
  })
  public sealed interface Plot
      permits Area,
          Area3D,
          Bar,
          Bar3D,
          Doughnut,
          Line,
          Line3D,
          Pie,
          Pie3D,
          Radar,
          Scatter,
          Surface,
          Surface3D {}

  /** Area chart plot. */
  public record Area(
      Boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    @JsonCreator
    static Area create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Area(
          defaultBoolean(varyColors),
          grouping == null ? ExcelChartGrouping.STANDARD : grouping,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Area {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one area plot with the default category/value axis pair. */
    public Area(Boolean varyColors, ExcelChartGrouping grouping, List<Series> series) {
      this(varyColors, grouping, defaultCategoryAxes(), series);
    }
  }

  /** 3D area chart plot. */
  public record Area3D(
      Boolean varyColors,
      ExcelChartGrouping grouping,
      Integer gapDepth,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    @JsonCreator
    static Area3D create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("gapDepth") Integer gapDepth,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Area3D(
          defaultBoolean(varyColors),
          grouping == null ? ExcelChartGrouping.STANDARD : grouping,
          gapDepth,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Area3D {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D area plot with the default category/value axis pair. */
    public Area3D(
        Boolean varyColors, ExcelChartGrouping grouping, Integer gapDepth, List<Series> series) {
      this(varyColors, grouping, gapDepth, defaultCategoryAxes(), series);
    }
  }

  /** Bar chart plot. */
  public record Bar(
      Boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Integer gapWidth,
      Integer overlap,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    @JsonCreator
    static Bar create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("barDirection") ExcelChartBarDirection barDirection,
        @JsonProperty("grouping") ExcelChartBarGrouping grouping,
        @JsonProperty("gapWidth") Integer gapWidth,
        @JsonProperty("overlap") Integer overlap,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Bar(
          defaultBoolean(varyColors),
          barDirection == null ? ExcelChartBarDirection.COLUMN : barDirection,
          grouping == null ? ExcelChartBarGrouping.CLUSTERED : grouping,
          gapWidth,
          overlap,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Bar {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
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
        Boolean varyColors,
        ExcelChartBarDirection barDirection,
        ExcelChartBarGrouping grouping,
        Integer gapWidth,
        Integer overlap,
        List<Series> series) {
      this(varyColors, barDirection, grouping, gapWidth, overlap, defaultCategoryAxes(), series);
    }
  }

  /** 3D bar chart plot. */
  public record Bar3D(
      Boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Integer gapDepth,
      Integer gapWidth,
      ExcelChartBarShape shape,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    @JsonCreator
    static Bar3D create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("barDirection") ExcelChartBarDirection barDirection,
        @JsonProperty("grouping") ExcelChartBarGrouping grouping,
        @JsonProperty("gapDepth") Integer gapDepth,
        @JsonProperty("gapWidth") Integer gapWidth,
        @JsonProperty("shape") ExcelChartBarShape shape,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Bar3D(
          defaultBoolean(varyColors),
          barDirection == null ? ExcelChartBarDirection.COLUMN : barDirection,
          grouping == null ? ExcelChartBarGrouping.CLUSTERED : grouping,
          gapDepth,
          gapWidth,
          shape,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Bar3D {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D bar plot with the default category/value axis pair. */
    public Bar3D(
        Boolean varyColors,
        ExcelChartBarDirection barDirection,
        ExcelChartBarGrouping grouping,
        Integer gapDepth,
        Integer gapWidth,
        ExcelChartBarShape shape,
        List<Series> series) {
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
  public record Doughnut(
      Boolean varyColors, Integer firstSliceAngle, Integer holeSize, List<Series> series)
      implements Plot {
    @JsonCreator
    static Doughnut create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("firstSliceAngle") Integer firstSliceAngle,
        @JsonProperty("holeSize") Integer holeSize,
        @JsonProperty("series") List<Series> series) {
      return new Doughnut(defaultBoolean(varyColors), firstSliceAngle, holeSize, series);
    }

    public Doughnut {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      validateAngle(firstSliceAngle);
      if (holeSize != null && (holeSize < 10 || holeSize > 90)) {
        throw new IllegalArgumentException("holeSize must be between 10 and 90");
      }
      series = copySeries(series);
    }
  }

  /** Line chart plot. */
  public record Line(
      Boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    @JsonCreator
    static Line create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Line(
          defaultBoolean(varyColors),
          grouping == null ? ExcelChartGrouping.STANDARD : grouping,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Line {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one line plot with the default category/value axis pair. */
    public Line(Boolean varyColors, ExcelChartGrouping grouping, List<Series> series) {
      this(varyColors, grouping, defaultCategoryAxes(), series);
    }
  }

  /** 3D line chart plot. */
  public record Line3D(
      Boolean varyColors,
      ExcelChartGrouping grouping,
      Integer gapDepth,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    @JsonCreator
    static Line3D create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("grouping") ExcelChartGrouping grouping,
        @JsonProperty("gapDepth") Integer gapDepth,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Line3D(
          defaultBoolean(varyColors),
          grouping == null ? ExcelChartGrouping.STANDARD : grouping,
          gapDepth,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Line3D {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D line plot with the default category/value axis pair. */
    public Line3D(
        Boolean varyColors, ExcelChartGrouping grouping, Integer gapDepth, List<Series> series) {
      this(varyColors, grouping, gapDepth, defaultCategoryAxes(), series);
    }
  }

  /** Pie chart plot. */
  public record Pie(Boolean varyColors, Integer firstSliceAngle, List<Series> series)
      implements Plot {
    @JsonCreator
    static Pie create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("firstSliceAngle") Integer firstSliceAngle,
        @JsonProperty("series") List<Series> series) {
      return new Pie(defaultBoolean(varyColors), firstSliceAngle, series);
    }

    public Pie {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      validateAngle(firstSliceAngle);
      series = copySeries(series);
    }
  }

  /** 3D pie chart plot. */
  public record Pie3D(Boolean varyColors, List<Series> series) implements Plot {
    @JsonCreator
    static Pie3D create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("series") List<Series> series) {
      return new Pie3D(defaultBoolean(varyColors), series);
    }

    public Pie3D {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      series = copySeries(series);
    }
  }

  /** Radar chart plot. */
  public record Radar(
      Boolean varyColors, ExcelChartRadarStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    @JsonCreator
    static Radar create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("style") ExcelChartRadarStyle style,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Radar(
          defaultBoolean(varyColors),
          style == null ? ExcelChartRadarStyle.STANDARD : style,
          axes == null ? defaultCategoryAxes() : axes,
          series);
    }

    public Radar {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(style, "style must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one radar plot with the default category/value axis pair. */
    public Radar(Boolean varyColors, ExcelChartRadarStyle style, List<Series> series) {
      this(varyColors, style, defaultCategoryAxes(), series);
    }
  }

  /** Scatter chart plot. */
  public record Scatter(
      Boolean varyColors, ExcelChartScatterStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    @JsonCreator
    static Scatter create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("style") ExcelChartScatterStyle style,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Scatter(
          defaultBoolean(varyColors),
          style == null ? ExcelChartScatterStyle.LINE_MARKER : style,
          axes == null ? defaultScatterAxes() : axes,
          series);
    }

    public Scatter {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(style, "style must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one scatter plot with the default X/Y axis pair. */
    public Scatter(Boolean varyColors, ExcelChartScatterStyle style, List<Series> series) {
      this(varyColors, style, defaultScatterAxes(), series);
    }
  }

  /** Surface chart plot. */
  public record Surface(Boolean varyColors, Boolean wireframe, List<Axis> axes, List<Series> series)
      implements Plot {
    @JsonCreator
    static Surface create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("wireframe") Boolean wireframe,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Surface(
          defaultBoolean(varyColors),
          defaultBoolean(wireframe),
          axes == null ? defaultSurfaceAxes() : axes,
          series);
    }

    public Surface {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(wireframe, "wireframe must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one surface plot with the default category/value/series axes. */
    public Surface(Boolean varyColors, Boolean wireframe, List<Series> series) {
      this(varyColors, wireframe, defaultSurfaceAxes(), series);
    }
  }

  /** 3D surface chart plot. */
  public record Surface3D(
      Boolean varyColors, Boolean wireframe, List<Axis> axes, List<Series> series) implements Plot {
    @JsonCreator
    static Surface3D create(
        @JsonProperty("varyColors") Boolean varyColors,
        @JsonProperty("wireframe") Boolean wireframe,
        @JsonProperty("axes") List<Axis> axes,
        @JsonProperty("series") List<Series> series) {
      return new Surface3D(
          defaultBoolean(varyColors),
          defaultBoolean(wireframe),
          axes == null ? defaultSurfaceAxes() : axes,
          series);
    }

    public Surface3D {
      Objects.requireNonNull(varyColors, "varyColors must not be null");
      Objects.requireNonNull(wireframe, "wireframe must not be null");
      axes = copyAxes(axes, "axes");
      series = copySeries(series);
    }

    /** Creates one 3D surface plot with the default category/value/series axes. */
    public Surface3D(Boolean varyColors, Boolean wireframe, List<Series> series) {
      this(varyColors, wireframe, defaultSurfaceAxes(), series);
    }
  }

  private static List<Axis> defaultCategoryAxes() {
    return List.of(
        new Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<Axis> defaultScatterAxes() {
    return List.of(
        new Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<Axis> defaultSurfaceAxes() {
    return List.of(
        new Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<Series> copySeries(List<Series> series) {
    return copyNonEmptyValues(series, "series");
  }

  private static List<Axis> copyAxes(List<Axis> axes, String fieldName) {
    return copyNonEmptyValues(axes, fieldName);
  }

  private static void validateAngle(Integer angle) {
    if (angle != null && (angle < 0 || angle > 360)) {
      throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
    }
  }

  private static boolean defaultBoolean(Boolean value) {
    return value != null && value;
  }

  private static <T> List<T> copyNonEmptyValues(List<T> values, String fieldName) {
    List<T> copiedValues = copyValues(values, fieldName);
    if (copiedValues.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copiedValues;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copiedValues = new ArrayList<>(values.size());
    for (T value : values) {
      copiedValues.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copiedValues);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
