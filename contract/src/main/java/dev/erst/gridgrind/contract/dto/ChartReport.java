package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual chart report returned by chart reads. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartReport.Bar.class, name = "BAR"),
  @JsonSubTypes.Type(value = ChartReport.Line.class, name = "LINE"),
  @JsonSubTypes.Type(value = ChartReport.Pie.class, name = "PIE"),
  @JsonSubTypes.Type(value = ChartReport.Unsupported.class, name = "UNSUPPORTED")
})
public sealed interface ChartReport
    permits ChartReport.Bar, ChartReport.Line, ChartReport.Pie, ChartReport.Unsupported {

  /** Sheet-local chart name. */
  String name();

  /** Stored anchor backing the chart frame. */
  DrawingAnchorReport anchor();

  /** Supported simple bar chart report. */
  record Bar(
      String name,
      DrawingAnchorReport anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      List<Axis> axes,
      List<Series> series)
      implements ChartReport {
    public Bar {
      validateCommon(name, anchor);
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      axes = copyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Supported simple line chart report. */
  record Line(
      String name,
      DrawingAnchorReport anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      List<Axis> axes,
      List<Series> series)
      implements ChartReport {
    public Line {
      validateCommon(name, anchor);
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      axes = copyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Supported simple pie chart report. */
  record Pie(
      String name,
      DrawingAnchorReport anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      Integer firstSliceAngle,
      List<Series> series)
      implements ChartReport {
    public Pie {
      validateCommon(name, anchor);
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      if (firstSliceAngle != null && (firstSliceAngle < 0 || firstSliceAngle > 360)) {
        throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
      }
      series = copyNonEmptyValues(series, "series");
    }
  }

  /** Factual unsupported chart report preserved without first-class mutation support. */
  record Unsupported(
      String name, DrawingAnchorReport anchor, List<String> plotTypeTokens, String detail)
      implements ChartReport {
    public Unsupported {
      validateCommon(name, anchor);
      plotTypeTokens = copyValues(plotTypeTokens, "plotTypeTokens");
      for (String plotTypeToken : plotTypeTokens) {
        requireNonBlank(plotTypeToken, "plotTypeTokens value");
      }
      detail = requireNonBlank(detail, "detail");
    }
  }

  /** Factual chart title. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Title.None.class, name = "NONE"),
    @JsonSubTypes.Type(value = Title.Text.class, name = "TEXT"),
    @JsonSubTypes.Type(value = Title.Formula.class, name = "FORMULA")
  })
  sealed interface Title permits Title.None, Title.Text, Title.Formula {

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
  sealed interface Legend permits Legend.Hidden, Legend.Visible {

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
  record Axis(
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
  record Series(Title title, DataSource categories, DataSource values) {
    public Series {
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(categories, "categories must not be null");
      Objects.requireNonNull(values, "values must not be null");
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
  sealed interface DataSource
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

  private static void validateCommon(String name, DrawingAnchorReport anchor) {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
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
