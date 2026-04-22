package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Authored chart definition used by workbook-core mutations. */
public record ExcelChartDefinition(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    Title title,
    Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    List<Plot> plots) {
  public ExcelChartDefinition {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    plots = copyNonEmptyValues(plots, "plots");
  }

  /** Supported authored chart title shapes. */
  public sealed interface Title permits Title.None, Title.Text, Title.Formula {

    /** Explicitly remove any chart or series title. */
    record None() implements Title {}

    /** Static text title. */
    record Text(String text) implements Title {
      public Text {
        Objects.requireNonNull(text, "text must not be null");
        if (text.isBlank()) {
          throw new IllegalArgumentException("text must not be blank");
        }
      }
    }

    /** Formula-backed title. */
    record Formula(String formula) implements Title {
      public Formula {
        formula = requireNonBlank(formula, "formula");
      }
    }
  }

  /** Supported authored legend states. */
  public sealed interface Legend permits Legend.Hidden, Legend.Visible {

    /** Hide the chart legend entirely. */
    record Hidden() implements Legend {}

    /** Show the legend at the requested position. */
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** Authored axis definition. */
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

    /** Convenience overload for authored series that only need title, categories, and values. */
    public Series(Title title, DataSource categories, DataSource values) {
      this(title, categories, values, null, null, null, null);
    }
  }

  /** Authored chart data source. */
  public sealed interface DataSource
      permits DataSource.Reference, DataSource.StringLiteral, DataSource.NumericLiteral {

    /** Workbook-bound chart source formula. */
    record Reference(String formula) implements DataSource {
      public Reference {
        formula = requireNonBlank(formula, "formula");
      }
    }

    /** Literal string source stored directly inside the chart part. */
    record StringLiteral(List<String> values) implements DataSource {
      public StringLiteral {
        values = copyValues(values, "values");
      }
    }

    /** Literal numeric source stored directly inside the chart part. */
    record NumericLiteral(List<Double> values) implements DataSource {
      public NumericLiteral {
        values = copyValues(values, "values");
      }
    }
  }

  /** Authored chart plot. */
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
