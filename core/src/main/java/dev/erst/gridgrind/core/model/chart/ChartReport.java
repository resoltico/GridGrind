package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual chart report returned by chart reads. */
@CatalogType(
    id = "ChartReport",
    summary = "One factual chart report with chart-level presentation state and one or more plots.",
    displayOrder = 108)
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
  @Union(discriminator = "type", catalog = true, displayOrder = 12)
  public sealed interface Title
      permits dev.erst.gridgrind.core.model.chart.ChartReport.Title.None,
          dev.erst.gridgrind.core.model.chart.ChartReport.Title.Text,
          dev.erst.gridgrind.core.model.chart.ChartReport.Title.Formula {

    /** No chart or series title is stored. */
    @Variant(id = "NONE", summary = "No title is present.")
    record None() implements Title {}

    /** Static title text. */
    @Variant(id = "TEXT", summary = "Static title text.")
    record Text(String text) implements Title {
      public Text {
        Objects.requireNonNull(text, "text must not be null");
      }
    }

    /** Formula-backed title with cached text. */
    @Variant(id = "FORMULA", summary = "Formula-backed title with cached text.")
    record Formula(String formula, String cachedText) implements Title {
      public Formula {
        formula = requireNonBlank(formula, "formula");
        Objects.requireNonNull(cachedText, "cachedText must not be null");
      }
    }
  }

  /** Factual legend state. */
  @Union(discriminator = "type", catalog = true, displayOrder = 13)
  public sealed interface Legend
      permits dev.erst.gridgrind.core.model.chart.ChartReport.Legend.Hidden,
          dev.erst.gridgrind.core.model.chart.ChartReport.Legend.Visible {

    /** No legend is stored. */
    @Variant(id = "HIDDEN", summary = "No legend is present.")
    record Hidden() implements Legend {}

    /** Visible legend at one position. */
    @Variant(id = "VISIBLE", summary = "Visible legend at one persisted position.")
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** Factual chart axis. */
  @CatalogType(id = "ChartAxisReport", summary = "Exact chart-axis report.", displayOrder = 109)
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
  @CatalogType(
      id = "ChartSeriesReport",
      summary =
          "Exact chart-series report. smooth, marker, and explosion fields are populated only when the stored plot family supports them.",
      displayOrder = 110)
  public record Series(
      Title title,
      DataSource categories,
      DataSource values,
      Optional<Boolean> smooth,
      Optional<ExcelChartMarkerStyle> markerStyle,
      Optional<Short> markerSize,
      Optional<Long> explosion) {
    public Series {
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(categories, "categories must not be null");
      Objects.requireNonNull(values, "values must not be null");
      Objects.requireNonNull(smooth, "smooth must not be null");
      Objects.requireNonNull(markerStyle, "markerStyle must not be null");
      Objects.requireNonNull(markerSize, "markerSize must not be null");
      Objects.requireNonNull(explosion, "explosion must not be null");
      if (markerSize.isPresent()
          && (markerSize.orElseThrow() < 2 || markerSize.orElseThrow() > 72)) {
        throw new IllegalArgumentException("markerSize must be between 2 and 72");
      }
      if (explosion.isPresent() && explosion.orElseThrow() < 0L) {
        throw new IllegalArgumentException("explosion must not be negative");
      }
    }
  }

  /** Factual chart data source. */
  @Union(discriminator = "type", catalog = true, displayOrder = 14)
  public sealed interface DataSource
      permits dev.erst.gridgrind.core.model.chart.ChartReport.DataSource.StringReference,
          dev.erst.gridgrind.core.model.chart.ChartReport.DataSource.NumericReference,
          dev.erst.gridgrind.core.model.chart.ChartReport.DataSource.StringLiteral,
          dev.erst.gridgrind.core.model.chart.ChartReport.DataSource.NumericLiteral {

    /** Formula-backed string source. */
    @Variant(
        id = "STRING_REFERENCE",
        summary = "Formula-backed string chart source plus cached values.")
    record StringReference(String formula, List<String> cachedValues) implements DataSource {
      public StringReference {
        formula = requireNonBlank(formula, "formula");
        cachedValues = copyValues(cachedValues, "cachedValues");
      }
    }

    /** Formula-backed numeric source. */
    @Variant(
        id = "NUMERIC_REFERENCE",
        summary = "Formula-backed numeric chart source plus cached values.")
    record NumericReference(String formula, Optional<String> formatCode, List<String> cachedValues)
        implements DataSource {
      public NumericReference {
        formula = requireNonBlank(formula, "formula");
        Objects.requireNonNull(formatCode, "formatCode must not be null");
        if (formatCode.isPresent() && formatCode.orElseThrow().isBlank()) {
          throw new IllegalArgumentException("formatCode must not be blank");
        }
        cachedValues = copyValues(cachedValues, "cachedValues");
      }
    }

    /** Literal string source stored directly in the chart part. */
    @Variant(
        id = "STRING_LITERAL",
        summary = "Literal string chart source stored directly in the chart part.")
    record StringLiteral(List<String> values) implements DataSource {
      public StringLiteral {
        values = copyValues(values, "values");
      }
    }

    /** Literal numeric source stored directly in the chart part. */
    @Variant(
        id = "NUMERIC_LITERAL",
        summary = "Literal numeric chart source stored directly in the chart part.")
    record NumericLiteral(Optional<String> formatCode, List<String> values) implements DataSource {
      public NumericLiteral {
        Objects.requireNonNull(formatCode, "formatCode must not be null");
        if (formatCode.isPresent() && formatCode.orElseThrow().isBlank()) {
          throw new IllegalArgumentException("formatCode must not be blank");
        }
        values = copyValues(values, "values");
      }
    }
  }

  /** One factual chart plot. */
  @Union(discriminator = "type", catalog = true, displayOrder = 11)
  public sealed interface Plot
      permits dev.erst.gridgrind.core.model.chart.ChartReport.Area,
          dev.erst.gridgrind.core.model.chart.ChartReport.Area3D,
          dev.erst.gridgrind.core.model.chart.ChartReport.Bar,
          dev.erst.gridgrind.core.model.chart.ChartReport.Bar3D,
          dev.erst.gridgrind.core.model.chart.ChartReport.Doughnut,
          dev.erst.gridgrind.core.model.chart.ChartReport.Line,
          dev.erst.gridgrind.core.model.chart.ChartReport.Line3D,
          dev.erst.gridgrind.core.model.chart.ChartReport.Pie,
          dev.erst.gridgrind.core.model.chart.ChartReport.Pie3D,
          dev.erst.gridgrind.core.model.chart.ChartReport.Radar,
          dev.erst.gridgrind.core.model.chart.ChartReport.Scatter,
          dev.erst.gridgrind.core.model.chart.ChartReport.Surface,
          dev.erst.gridgrind.core.model.chart.ChartReport.Surface3D,
          dev.erst.gridgrind.core.model.chart.ChartReport.Unsupported {}

  /** Factual area chart plot. */
  @Variant(id = "AREA", summary = "Exact area-chart plot report.")
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
  @Variant(id = "AREA_3D", summary = "Exact 3D area-chart plot report.")
  public record Area3D(
      boolean varyColors,
      ExcelChartGrouping grouping,
      Optional<Integer> gapDepth,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Area3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual bar chart plot. */
  @Variant(id = "BAR", summary = "Exact bar-chart plot report.")
  public record Bar(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapWidth,
      Optional<Integer> overlap,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Bar {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapWidth, "gapWidth must not be null");
      Objects.requireNonNull(overlap, "overlap must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D bar chart plot. */
  @Variant(id = "BAR_3D", summary = "Exact 3D bar-chart plot report.")
  public record Bar3D(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapDepth,
      Optional<Integer> gapWidth,
      Optional<ExcelChartBarShape> shape,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Bar3D {
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      Objects.requireNonNull(gapWidth, "gapWidth must not be null");
      Objects.requireNonNull(shape, "shape must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual doughnut chart plot. */
  @Variant(id = "DOUGHNUT", summary = "Exact doughnut-chart plot report.")
  public record Doughnut(
      boolean varyColors,
      Optional<Integer> firstSliceAngle,
      Optional<Integer> holeSize,
      List<Series> series)
      implements Plot {
    public Doughnut {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      Objects.requireNonNull(holeSize, "holeSize must not be null");
      validateAngle(firstSliceAngle);
      if (holeSize.isPresent() && (holeSize.orElseThrow() < 10 || holeSize.orElseThrow() > 90)) {
        throw new IllegalArgumentException("holeSize must be between 10 and 90");
      }
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual line chart plot. */
  @Variant(id = "LINE", summary = "Exact line-chart plot report.")
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
  @Variant(id = "LINE_3D", summary = "Exact 3D line-chart plot report.")
  public record Line3D(
      boolean varyColors,
      ExcelChartGrouping grouping,
      Optional<Integer> gapDepth,
      List<Axis> axes,
      List<Series> series)
      implements Plot {
    public Line3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual pie chart plot. */
  @Variant(id = "PIE", summary = "Exact pie-chart plot report.")
  public record Pie(boolean varyColors, Optional<Integer> firstSliceAngle, List<Series> series)
      implements Plot {
    public Pie {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      validateAngle(firstSliceAngle);
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D pie chart plot. */
  @Variant(id = "PIE_3D", summary = "Exact 3D pie-chart plot report.")
  public record Pie3D(boolean varyColors, List<Series> series) implements Plot {
    public Pie3D {
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual radar chart plot. */
  @Variant(id = "RADAR", summary = "Exact radar-chart plot report.")
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
  @Variant(id = "SCATTER", summary = "Exact scatter-chart plot report.")
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
  @Variant(id = "SURFACE", summary = "Exact surface-chart plot report.")
  public record Surface(boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series)
      implements Plot {
    public Surface {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual 3D surface chart plot. */
  @Variant(id = "SURFACE_3D", summary = "Exact 3D surface-chart plot report.")
  public record Surface3D(
      boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series) implements Plot {
    public Surface3D {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Plot fact preserved without a first-class typed model. */
  @Variant(
      id = "UNSUPPORTED",
      summary = "Exact unsupported chart plot report preserved from the workbook.")
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

  private static void validateAngle(Optional<Integer> firstSliceAngle) {
    if (firstSliceAngle.isPresent()
        && (firstSliceAngle.orElseThrow() < 0 || firstSliceAngle.orElseThrow() > 360)) {
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
