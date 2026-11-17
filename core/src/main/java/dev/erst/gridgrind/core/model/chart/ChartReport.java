package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.FieldRule;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
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
import dev.erst.gridgrind.core.vocabulary.TextValidation;
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
  private static final long MAX_UNSIGNED_DECLARATION = 0xffff_ffffL;

  public ChartReport {
    Objects.requireNonNull(name, "name must not be null");
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

    /** Formula-backed title and its actual stored string-cache facts. */
    @Variant(
        id = "FORMULA",
        summary = "Stored formula-backed title with its cache; never live-resolved text.")
    record Formula(Optional<String> formula, StringCache cache) implements Title {
      public Formula {
        Objects.requireNonNull(formula, "formula must not be null");
        Objects.requireNonNull(cache, "cache must not be null");
      }
    }
  }

  /** Stored string-cache presence and physical point order, independent of live sources. */
  @Union(discriminator = "type", catalog = true, displayOrder = 15)
  public sealed interface StringCache permits StringCache.Missing, StringCache.Stored {
    /** No string cache element is stored. */
    @Variant(id = "MISSING", summary = "No string cache element is stored.")
    record Missing() implements StringCache {}

    /** Declared count and points exactly as stored, without padding or sorting. */
    @Variant(
        id = "STORED",
        summary = "Stored string cache with its optional declared count and physical point order.")
    record Stored(
        @FieldRule(
                kind = FieldRule.Kind.NUMBER_RANGE,
                minimum = 0,
                maximum = MAX_UNSIGNED_DECLARATION)
            Optional<Long> declaredCount,
        List<StringPoint> points)
        implements StringCache {
      public Stored {
        declaredCount = unsignedDeclaration(declaredCount, "declaredCount");
        points = copyValues(points, "points");
      }
    }
  }

  /** One stored string point; absent value and stored empty text remain distinct. */
  @CatalogType(
      id = "ChartStringCachePoint",
      summary =
          "One stored string-cache point. Missing index/value declarations stay absent; list"
              + " position does not infer the index.",
      displayOrder = 111)
  public record StringPoint(
      @FieldRule(
              kind = FieldRule.Kind.NUMBER_RANGE,
              minimum = 0,
              maximum = MAX_UNSIGNED_DECLARATION)
          Optional<Long> index,
      Optional<String> valueText) {
    public StringPoint {
      index = unsignedDeclaration(index, "index");
      Objects.requireNonNull(valueText, "valueText must not be null");
    }
  }

  /** Stored numeric-cache facts; point values retain raw text, including errors and blanks. */
  @Union(discriminator = "type", catalog = true, displayOrder = 16)
  public sealed interface NumericCache permits NumericCache.Missing, NumericCache.Stored {
    /** No numeric cache element is stored. */
    @Variant(id = "MISSING", summary = "No numeric cache element is stored.")
    record Missing() implements NumericCache {}

    /** Stored numeric cache with its declared count, formats and physical point order. */
    @Variant(
        id = "STORED",
        summary =
            "Stored numeric cache. Count, formats and raw point text are never inferred from live"
                + " cells.")
    record Stored(
        @FieldRule(
                kind = FieldRule.Kind.NUMBER_RANGE,
                minimum = 0,
                maximum = MAX_UNSIGNED_DECLARATION)
            Optional<Long> declaredCount,
        Optional<String> formatCode,
        List<NumericPoint> points)
        implements NumericCache {
      public Stored {
        declaredCount = unsignedDeclaration(declaredCount, "declaredCount");
        Objects.requireNonNull(formatCode, "formatCode must not be null");
        points = copyValues(points, "points");
      }
    }
  }

  /** One stored numeric point; no finite-number or authored-value conversion is applied. */
  @CatalogType(
      id = "ChartNumericCachePoint",
      summary =
          "One stored numeric-cache point with raw value text and optional per-point format."
              + " Missing, blank, error and nonfinite-looking text remain distinct facts.",
      displayOrder = 112)
  public record NumericPoint(
      @FieldRule(
              kind = FieldRule.Kind.NUMBER_RANGE,
              minimum = 0,
              maximum = MAX_UNSIGNED_DECLARATION)
          Optional<Long> index,
      Optional<String> valueText,
      Optional<String> formatCode) {
    public NumericPoint {
      index = unsignedDeclaration(index, "index");
      Objects.requireNonNull(valueText, "valueText must not be null");
      Objects.requireNonNull(formatCode, "formatCode must not be null");
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

    /** Visible legend with its effective native position, including the default. */
    @Variant(
        id = "VISIBLE",
        summary = "Visible legend with its effective native position, including the default.")
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** Stored axis crossing choice, without an authored default or numeric conversion. */
  @Union(discriminator = "type", catalog = true, displayOrder = 17)
  public sealed interface AxisCrossing
      permits AxisCrossing.Missing, AxisCrossing.Mode, AxisCrossing.Value {
    /** Neither optional native crossing declaration is present. */
    @Variant(id = "MISSING", summary = "No axis crossing declaration is stored.")
    record Missing() implements AxisCrossing {}

    /** Stored native crossing mode. */
    @Variant(id = "MODE", summary = "Stored axis crossing mode.")
    record Mode(ExcelChartAxisCrosses mode) implements AxisCrossing {
      /** Retains the stored mode without applying an authored default. */
      public Mode {
        Objects.requireNonNull(mode, "mode must not be null");
      }
    }

    /**
     * Exact validated native xsd:double attribute text, including exponent and nonfinite values.
     */
    @Variant(
        id = "VALUE",
        summary = "Stored numeric axis crossing with its exact native attribute text.")
    record Value(String valueText) implements AxisCrossing {
      /** Retains validated native numeric text without parsing or canonicalizing it. */
      public Value {
        Objects.requireNonNull(valueText, "valueText must not be null");
      }
    }
  }

  /** Factual chart axis. */
  @CatalogType(id = "ChartAxisReport", summary = "Exact chart-axis report.", displayOrder = 109)
  public record Axis(
      ExcelChartAxisKind kind,
      ExcelChartAxisPosition position,
      AxisCrossing crosses,
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
          "Exact chart-series report. smooth, marker, and explosion fields are populated only when"
              + " the stored plot family supports them.",
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

  /** Factual chart data source; every cache is read from stored chart XML only. */
  @Union(discriminator = "type", catalog = true, displayOrder = 14)
  public sealed interface DataSource
      permits DataSource.StringReference,
          DataSource.NumericReference,
          DataSource.StringLiteral,
          DataSource.NumericLiteral,
          DataSource.Missing,
          DataSource.Unsupported {
    /** Formula-backed string source without live reference resolution. */
    @Variant(
        id = "STRING_REFERENCE",
        summary =
            "Stored string reference and string-cache facts; formula absence/blankness is"
                + " preserved.")
    record StringReference(Optional<String> formula, StringCache cache) implements DataSource {
      public StringReference {
        Objects.requireNonNull(formula, "formula must not be null");
        Objects.requireNonNull(cache, "cache must not be null");
      }
    }

    /** Formula-backed numeric source without live reference resolution. */
    @Variant(
        id = "NUMERIC_REFERENCE",
        summary =
            "Stored numeric reference and numeric-cache facts; no values are substituted from live"
                + " cells.")
    record NumericReference(Optional<String> formula, NumericCache cache) implements DataSource {
      public NumericReference {
        Objects.requireNonNull(formula, "formula must not be null");
        Objects.requireNonNull(cache, "cache must not be null");
      }
    }

    /** Literal string payload stored directly in the chart. */
    @Variant(
        id = "STRING_LITERAL",
        summary = "Stored literal string data, including declared count and physical point order.")
    record StringLiteral(StringCache.Stored data) implements DataSource {
      public StringLiteral {
        Objects.requireNonNull(data, "data must not be null");
      }
    }

    /** Literal numeric payload stored directly in the chart. */
    @Variant(
        id = "NUMERIC_LITERAL",
        summary = "Stored literal numeric data, retaining raw value text, counts and formats.")
    record NumericLiteral(NumericCache.Stored data) implements DataSource {
      public NumericLiteral {
        Objects.requireNonNull(data, "data must not be null");
      }
    }

    /** No supported data-source element is stored. */
    @Variant(id = "MISSING", summary = "No chart data-source element is stored.")
    record Missing() implements DataSource {}

    /** A stored data-source kind outside the typed observation contract. */
    @Variant(
        id = "UNSUPPORTED",
        summary =
            "A stored chart data-source kind outside the typed observation contract; no substitute"
                + " values are reported.")
    record Unsupported(String nativeType, String detail) implements DataSource {
      public Unsupported {
        nativeType = TextValidation.requireNonBlank(nativeType, "nativeType");
        detail = TextValidation.requireNonBlank(detail, "detail");
      }

      /** Describes an unmodeled stored data-source kind without substituting values. */
      public static Unsupported unmodeled(String nativeType) {
        return new Unsupported(
            nativeType,
            "Stored chart data-source kind is outside the modeled observation contract.");
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
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
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
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
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
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
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
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
    }
  }

  /** Factual doughnut chart plot. */
  @Variant(id = "DOUGHNUT", summary = "Exact doughnut-chart plot report.")
  public record Doughnut(
      boolean varyColors,
      Optional<Integer> firstSliceAngle,
      @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 1, maximum = 90)
          Optional<Integer> holeSize,
      List<Series> series)
      implements Plot {
    public Doughnut {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      Objects.requireNonNull(holeSize, "holeSize must not be null");
      validateAngle(firstSliceAngle);
      if (holeSize.isPresent() && (holeSize.orElseThrow() < 1 || holeSize.orElseThrow() > 90)) {
        throw new ValueInvariantException(
            new MessageInvariant("holeSize must be between 1 and 90", Optional.of("holeSize")));
      }
      series = copyValues(series, "series");
    }
  }

  /** Factual line chart plot. */
  @Variant(id = "LINE", summary = "Exact line-chart plot report.")
  public record Line(
      boolean varyColors, ExcelChartGrouping grouping, List<Axis> axes, List<Series> series)
      implements Plot {
    public Line {
      Objects.requireNonNull(grouping, "grouping must not be null");
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
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
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
    }
  }

  /** Factual pie chart plot. */
  @Variant(id = "PIE", summary = "Exact pie-chart plot report.")
  public record Pie(boolean varyColors, Optional<Integer> firstSliceAngle, List<Series> series)
      implements Plot {
    public Pie {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      validateAngle(firstSliceAngle);
      series = copyValues(series, "series");
    }
  }

  /** Factual 3D pie chart plot. */
  @Variant(id = "PIE_3D", summary = "Exact 3D pie-chart plot report.")
  public record Pie3D(boolean varyColors, List<Series> series) implements Plot {
    public Pie3D {
      series = copyValues(series, "series");
    }
  }

  /** Factual radar chart plot. */
  @Variant(id = "RADAR", summary = "Exact radar-chart plot report.")
  public record Radar(
      boolean varyColors, ExcelChartRadarStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Radar {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
    }
  }

  /** Factual scatter chart plot. */
  @Variant(id = "SCATTER", summary = "Exact scatter-chart plot report.")
  public record Scatter(
      boolean varyColors, ExcelChartScatterStyle style, List<Axis> axes, List<Series> series)
      implements Plot {
    public Scatter {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
    }
  }

  /** Factual surface chart plot. */
  @Variant(id = "SURFACE", summary = "Exact surface-chart plot report.")
  public record Surface(boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series)
      implements Plot {
    public Surface {
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
    }
  }

  /** Factual 3D surface chart plot. */
  @Variant(id = "SURFACE_3D", summary = "Exact 3D surface-chart plot report.")
  public record Surface3D(
      boolean varyColors, boolean wireframe, List<Axis> axes, List<Series> series) implements Plot {
    public Surface3D {
      axes = copyValues(axes, "axes");
      series = copyValues(series, "series");
    }
  }

  /** Plot fact preserved without a first-class typed model. */
  @Variant(
      id = "UNSUPPORTED",
      summary = "Exact unsupported chart plot report preserved from the workbook.")
  public record Unsupported(String plotTypeToken, String detail) implements Plot {
    public Unsupported {
      plotTypeToken = TextValidation.requireNonBlank(plotTypeToken, "plotTypeToken");
      detail = TextValidation.requireNonBlank(detail, "detail");
    }
  }

  private static Optional<Long> unsignedDeclaration(Optional<Long> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.filter(number -> number < 0 || number > MAX_UNSIGNED_DECLARATION).isPresent()) {
      throw new ValueInvariantException(
          new MessageInvariant(
              fieldName + " must be an unsigned 32-bit declaration", Optional.of(fieldName)));
    }
    return value;
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
}
