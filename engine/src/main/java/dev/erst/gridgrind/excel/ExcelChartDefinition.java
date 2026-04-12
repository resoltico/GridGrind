package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Supported authored chart definitions. */
public sealed interface ExcelChartDefinition
    permits ExcelChartDefinition.Bar, ExcelChartDefinition.Line, ExcelChartDefinition.Pie {

  /** Sheet-local chart name. */
  String name();

  /** Stored two-cell drawing anchor for the chart frame. */
  ExcelDrawingAnchor.TwoCell anchor();

  /** Authored chart title definition. */
  Title title();

  /** Authored legend visibility and position. */
  Legend legend();

  /** Blank-cell display behavior. */
  ExcelChartDisplayBlanksAs displayBlanksAs();

  /** Whether the chart should ignore hidden cells. */
  boolean plotOnlyVisibleCells();

  /** Whether the chart varies colors automatically. */
  boolean varyColors();

  /** Authored chart series in display order. */
  List<Series> series();

  /** Authored simple bar chart. */
  record Bar(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      List<Series> series)
      implements ExcelChartDefinition {
    public Bar {
      requireName(name);
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      Objects.requireNonNull(barDirection, "barDirection must not be null");
      series = copySeries(series);
    }
  }

  /** Authored simple line chart. */
  record Line(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      List<Series> series)
      implements ExcelChartDefinition {
    public Line {
      requireName(name);
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      series = copySeries(series);
    }
  }

  /** Authored simple pie chart. */
  record Pie(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      Integer firstSliceAngle,
      List<Series> series)
      implements ExcelChartDefinition {
    public Pie {
      requireName(name);
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(legend, "legend must not be null");
      Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
      if (firstSliceAngle != null && (firstSliceAngle < 0 || firstSliceAngle > 360)) {
        throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
      }
      series = copySeries(series);
    }
  }

  /** Supported authored chart title shapes. */
  sealed interface Title permits Title.None, Title.Text, Title.Formula {

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
  sealed interface Legend permits Legend.Hidden, Legend.Visible {

    /** Hide the chart legend entirely. */
    record Hidden() implements Legend {}

    /** Show the legend at the requested position. */
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** One authored chart series with workbook-bound category and value sources. */
  record Series(Title title, DataSource categories, DataSource values) {
    public Series {
      title = title == null ? new Title.None() : title;
      Objects.requireNonNull(categories, "categories must not be null");
      Objects.requireNonNull(values, "values must not be null");
    }
  }

  /** Workbook-bound chart source formula. */
  record DataSource(String formula) {
    public DataSource {
      formula = requireNonBlank(formula, "formula");
    }
  }

  private static List<Series> copySeries(List<Series> series) {
    Objects.requireNonNull(series, "series must not be null");
    List<Series> copiedSeries = List.copyOf(series);
    if (copiedSeries.isEmpty()) {
      throw new IllegalArgumentException("series must not be empty");
    }
    for (Series value : copiedSeries) {
      Objects.requireNonNull(value, "series must not contain null values");
    }
    return copiedSeries;
  }

  private static void requireName(String name) {
    requireNonBlank(name, "name");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
