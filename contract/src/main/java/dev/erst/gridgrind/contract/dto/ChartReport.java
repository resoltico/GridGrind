package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
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

/** Factual chart report returned by chart reads. */
public record ChartReport(
    String name,
    DrawingAnchorReport anchor,
    Title title,
    Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    List<Plot> plots) {
  public ChartReport {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    plots = copyNonEmptyValues(plots, "plots");
  }

  /** Factual chart title. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Title.None.class, name = "NONE"),
    @JsonSubTypes.Type(value = Title.Text.class, name = "TEXT"),
    @JsonSubTypes.Type(value = Title.Formula.class, name = "FORMULA")
  })
  public sealed interface Title permits Title.None, Title.Text, Title.Formula {

    /** No chart or series title is stored. */
    record None() implements Title {}

    /** Static title text. */
    record Text(String text) implements Title {
      public Text {
        Objects.requireNonNull(text, "text must not be null");
      }
    }

    /** Formula-backed title with cached text. */
    record Formula(String formula, String cachedText) implements Title {
      public Formula {
        formula = requireNonBlank(formula, "formula");
        Objects.requireNonNull(cachedText, "cachedText must not be null");
      }
    }
  }

  /** Factual legend state. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Legend.Hidden.class, name = "HIDDEN"),
    @JsonSubTypes.Type(value = Legend.Visible.class, name = "VISIBLE")
  })
  public sealed interface Legend permits Legend.Hidden, Legend.Visible {

    /** No legend is stored. */
    record Hidden() implements Legend {}

    /** Visible legend at one position. */
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** Factual chart axis. */
  public record Axis(
      ExcelChartAxisKind kind,
      ExcelChartAxisPosition position,
      ExcelChartAxisCrosses crosses,
      boolean visible) {
    public Axis {
      Objects.requireNonNull(kind, "kind must not be null");
      Objects.requireNonNull(position, "position must not be null");
      Objects.requireNonNull(crosses, "crosses must not be null");
    }
  }

  /** Factual chart series. */
  public record Series(
      Title title,
      DataSource categories,
      DataSource values,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      Long explosion) {
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
  }

  /** Factual chart data source. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = DataSource.StringReference.class, name = "STRING_REFERENCE"),
    @JsonSubTypes.Type(value = DataSource.NumericReference.class, name = "NUMERIC_REFERENCE"),
    @JsonSubTypes.Type(value = DataSource.StringLiteral.class, name = "STRING_LITERAL"),
    @JsonSubTypes.Type(value = DataSource.NumericLiteral.class, name = "NUMERIC_LITERAL")
  })
  public sealed interface DataSource
      permits DataSource.StringReference,
          DataSource.NumericReference,
          DataSource.StringLiteral,
          DataSource.NumericLiteral {

    /** Formula-backed string source. */
    record StringReference(String formula, List<String> cachedValues) implements DataSource {
      public StringReference {
        formula = requireNonBlank(formula, "formula");
        cachedValues = copyValues(cachedValues, "cachedValues");
      }
    }

    /** Formula-backed numeric source. */
    record NumericReference(String formula, String formatCode, List<String> cachedValues)
        implements DataSource {
      public NumericReference {
        formula = requireNonBlank(formula, "formula");
        if (formatCode != null && formatCode.isBlank()) {
          throw new IllegalArgumentException("formatCode must not be blank");
        }
        cachedValues = copyValues(cachedValues, "cachedValues");
      }
    }

    /** Literal string source stored directly in the chart part. */
    record StringLiteral(List<String> values) implements DataSource {
      public StringLiteral {
        values = copyValues(values, "values");
      }
    }

    /** Literal numeric source stored directly in the chart part. */
    record NumericLiteral(String formatCode, List<String> values) implements DataSource {
      public NumericLiteral {
        if (formatCode != null && formatCode.isBlank()) {
          throw new IllegalArgumentException("formatCode must not be blank");
        }
        values = copyValues(values, "values");
      }
    }
  }

  /** One factual chart plot. */
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
    @JsonSubTypes.Type(value = Surface3D.class, name = "SURFACE_3D"),
    @JsonSubTypes.Type(value = Unsupported.class, name = "UNSUPPORTED")
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
          Surface3D,
          Unsupported {}

  /** Factual area chart plot. */
  public record Area(
      boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    public Area {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D area chart plot. */
  public record Area3D(
      boolean varyColors,
      ExcelChartGrouping grouping,
      Integer gapDepth,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Area3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual bar chart plot. */
  public record Bar(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Integer gapWidth,
      Integer overlap,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Bar {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D bar chart plot. */
  public record Bar3D(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Integer gapDepth,
      Integer gapWidth,
      ExcelChartBarShape shape,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Bar3D {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual doughnut chart plot. */
  public record Doughnut(
      boolean varyColors, Integer firstSliceAngle, Integer holeSize, List<Series> series)
      implements Plot {
    public Doughnut {
      validateAngle(firstSliceAngle);
      if (holeSize != null && (holeSize < 10 || holeSize > 90)) {
        throw new IllegalArgumentException("holeSize must be between 10 and 90");
      }
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual line chart plot. */
  public record Line(
      boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    public Line {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D line chart plot. */
  public record Line3D(
      boolean varyColors,
      ExcelChartGrouping grouping,
      Integer gapDepth,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Line3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual pie chart plot. */
  public record Pie(boolean varyColors, Integer firstSliceAngle, List<Series> series)
      implements Plot {
    public Pie {
      validateAngle(firstSliceAngle);
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D pie chart plot. */
  public record Pie3D(boolean varyColors, List<Series> series) implements Plot {
    public Pie3D {
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual radar chart plot. */
  public record Radar(
      boolean varyColors, ExcelChartRadarStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Radar {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual scatter chart plot. */
  public record Scatter(
      boolean varyColors, ExcelChartScatterStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Scatter {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual surface chart plot. */
  public record Surface(boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series)
      implements Plot {
    public Surface {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D surface chart plot. */
  public record Surface3D(
      boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series) implements Plot {
    public Surface3D {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Plot fact preserved without a first-class typed model. */
  public record Unsupported(String plotTypeToken, String detail) implements Plot {
    public Unsupported {
      plotTypeToken = requireNonBlank(plotTypeToken, "plotTypeToken");
      detail = requireNonBlank(detail, "detail");
    }
  }

  private static <T> List<T> copyNonEmptyValues(List<T> values, String fieldName) {
    List<T> copiedValues = copyValues(values, fieldName);
    if (copiedValues.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copiedValues;
  }

  private static void validateAngle(Integer firstSliceAngle) {
    if (firstSliceAngle != null && (firstSliceAngle < 0 || firstSliceAngle > 360)) {
      throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
    }
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
