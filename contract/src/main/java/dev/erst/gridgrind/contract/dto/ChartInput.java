package dev.erst.gridgrind.contract.dto;

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
  public ChartInput {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    title = title == null ? new Title.None() : title;
    legend = legend == null ? new Legend.Visible(ExcelChartLegendPosition.RIGHT) : legend;
    displayBlanksAs = displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs;
    plotOnlyVisibleCells = plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells;
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
    public Axis {
      Objects.requireNonNull(kind, "kind must not be null");
      Objects.requireNonNull(position, "position must not be null");
      Objects.requireNonNull(crosses, "crosses must not be null");
      visible = visible == null ? Boolean.TRUE : visible;
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
    public Series {
      title = title == null ? new Title.None() : title;
      Objects.requireNonNull(categories, "categories must not be null");
      Objects.requireNonNull(values, "values must not be null");
      if (markerSize != null && (markerSize < 2 || markerSize > 72)) {
        throw new IllegalArgumentException("markerSize must be between 2 and 72");
      }
      if (explosion != null && explosion < 0L) {
        throw new IllegalArgumentException("explosion must not be negative");
      }
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
    public Area {
      varyColors = defaultBoolean(varyColors);
      grouping = grouping == null ? ExcelChartGrouping.STANDARD : grouping;
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
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
    public Area3D {
      varyColors = defaultBoolean(varyColors);
      grouping = grouping == null ? ExcelChartGrouping.STANDARD : grouping;
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
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
    public Bar {
      varyColors = defaultBoolean(varyColors);
      barDirection = barDirection == null ? ExcelChartBarDirection.COLUMN : barDirection;
      grouping = grouping == null ? ExcelChartBarGrouping.CLUSTERED : grouping;
      if (overlap != null && (overlap < -100 || overlap > 100)) {
        throw new IllegalArgumentException("overlap must be between -100 and 100");
      }
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
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
    public Bar3D {
      varyColors = defaultBoolean(varyColors);
      barDirection = barDirection == null ? ExcelChartBarDirection.COLUMN : barDirection;
      grouping = grouping == null ? ExcelChartBarGrouping.CLUSTERED : grouping;
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
    }
  }

  /** Doughnut chart plot. */
  public record Doughnut(
      Boolean varyColors, Integer firstSliceAngle, Integer holeSize, List<Series> series)
      implements Plot {
    public Doughnut {
      varyColors = defaultBoolean(varyColors);
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
    public Line {
      varyColors = defaultBoolean(varyColors);
      grouping = grouping == null ? ExcelChartGrouping.STANDARD : grouping;
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
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
    public Line3D {
      varyColors = defaultBoolean(varyColors);
      grouping = grouping == null ? ExcelChartGrouping.STANDARD : grouping;
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
    }
  }

  /** Pie chart plot. */
  public record Pie(Boolean varyColors, Integer firstSliceAngle, List<Series> series)
      implements Plot {
    public Pie {
      varyColors = defaultBoolean(varyColors);
      validateAngle(firstSliceAngle);
      series = copySeries(series);
    }
  }

  /** 3D pie chart plot. */
  public record Pie3D(Boolean varyColors, List<Series> series) implements Plot {
    public Pie3D {
      varyColors = defaultBoolean(varyColors);
      series = copySeries(series);
    }
  }

  /** Radar chart plot. */
  public record Radar(
      Boolean varyColors, ExcelChartRadarStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Radar {
      varyColors = defaultBoolean(varyColors);
      style = style == null ? ExcelChartRadarStyle.STANDARD : style;
      axes = copyAxes(axes, "axes", defaultCategoryAxes());
      series = copySeries(series);
    }
  }

  /** Scatter chart plot. */
  public record Scatter(
      Boolean varyColors, ExcelChartScatterStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Scatter {
      varyColors = defaultBoolean(varyColors);
      style = style == null ? ExcelChartScatterStyle.LINE_MARKER : style;
      axes = copyAxes(axes, "axes", defaultScatterAxes());
      series = copySeries(series);
    }
  }

  /** Surface chart plot. */
  public record Surface(Boolean varyColors, Boolean wireframe, List<Axis> axes, List<Series> series)
      implements Plot {
    public Surface {
      varyColors = defaultBoolean(varyColors);
      wireframe = defaultBoolean(wireframe);
      axes = copyAxes(axes, "axes", defaultSurfaceAxes());
      series = copySeries(series);
    }
  }

  /** 3D surface chart plot. */
  public record Surface3D(
      Boolean varyColors, Boolean wireframe, List<Axis> axes, List<Series> series) implements Plot {
    public Surface3D {
      varyColors = defaultBoolean(varyColors);
      wireframe = defaultBoolean(wireframe);
      axes = copyAxes(axes, "axes", defaultSurfaceAxes());
      series = copySeries(series);
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

  private static List<Axis> copyAxes(List<Axis> axes, String fieldName, List<Axis> defaults) {
    if (axes == null) {
      return defaults;
    }
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
