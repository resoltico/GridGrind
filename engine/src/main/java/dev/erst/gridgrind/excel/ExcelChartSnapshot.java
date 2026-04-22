package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Immutable factual chart snapshot returned by workbook reads. */
public record ExcelChartSnapshot(
    String name,
    ExcelDrawingAnchor anchor,
    Title title,
    Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    List<Plot> plots) {
  public ExcelChartSnapshot {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    plots = copyNonEmptyValues(plots, "plots");
  }

  /** Factual chart title shape. */
  public sealed interface Title permits Title.None, Title.Text, Title.Formula {

    /** No chart title is stored. */
    record None() implements Title {}

    /** Static title text. */
    record Text(String text) implements Title {
      public Text {
        Objects.requireNonNull(text, "text must not be null");
      }
    }

    /** Formula-backed title with the cached resolved text when present. */
    record Formula(String formula, String cachedText) implements Title {
      public Formula {
        formula = requireNonBlank(formula, "formula");
        Objects.requireNonNull(cachedText, "cachedText must not be null");
      }
    }
  }

  /** Factual legend state. */
  public sealed interface Legend permits Legend.Hidden, Legend.Visible {

    /** No chart legend is stored. */
    record Hidden() implements Legend {}

    /** Visible legend stored at the requested position. */
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

    /** Convenience overload for factual series that only need title, categories, and values. */
    public Series(Title title, DataSource categories, DataSource values) {
      this(title, categories, values, null, null, null, null);
    }
  }

  /** Factual chart data source. */
  public sealed interface DataSource
      permits DataSource.StringReference,
          DataSource.NumericReference,
          DataSource.StringLiteral,
          DataSource.NumericLiteral {

    /** Formula-backed string source with cached points. */
    record StringReference(String formula, List<String> cachedValues) implements DataSource {
      public StringReference {
        formula = requireNonBlank(formula, "formula");
        cachedValues = copyValues(cachedValues, "cachedValues");
      }
    }

    /** Formula-backed numeric source with cached points. */
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

    /** Literal string source stored directly inside the chart part. */
    record StringLiteral(List<String> values) implements DataSource {
      public StringLiteral {
        values = copyValues(values, "values");
      }
    }

    /** Literal numeric source stored directly inside the chart part. */
    record NumericLiteral(String formatCode, List<String> values) implements DataSource {
      public NumericLiteral {
        if (formatCode != null && formatCode.isBlank()) {
          throw new IllegalArgumentException("formatCode must not be blank");
        }
        values = copyValues(values, "values");
      }
    }
  }

  /** Factual chart plot. */
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

  public record Area(
      boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    public Area {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

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

  public record Doughnut(
      boolean varyColors, Integer firstSliceAngle, Integer holeSize, List<Series> series)
      implements Plot {
    public Doughnut {
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Line(
      boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    public Line {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

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

  public record Pie(boolean varyColors, Integer firstSliceAngle, List<Series> series)
      implements Plot {
    public Pie {
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Pie3D(boolean varyColors, List<Series> series) implements Plot {
    public Pie3D {
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Radar(
      boolean varyColors, ExcelChartRadarStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Radar {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Scatter(
      boolean varyColors, ExcelChartScatterStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Scatter {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Surface(boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series)
      implements Plot {
    public Surface {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Surface3D(
      boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series) implements Plot {
    public Surface3D {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

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
