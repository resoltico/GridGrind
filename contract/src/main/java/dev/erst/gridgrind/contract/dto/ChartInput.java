package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import java.util.List;
import java.util.Objects;

/** Supported authored chart inputs. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartInput.Bar.class, name = "BAR"),
  @JsonSubTypes.Type(value = ChartInput.Line.class, name = "LINE"),
  @JsonSubTypes.Type(value = ChartInput.Pie.class, name = "PIE")
})
public sealed interface ChartInput permits ChartInput.Bar, ChartInput.Line, ChartInput.Pie {

  /** Sheet-local chart name. */
  String name();

  /** Stored two-cell drawing anchor for the chart frame. */
  DrawingAnchorInput anchor();

  /** Requested chart title definition. */
  Title title();

  /** Requested legend visibility and position. */
  Legend legend();

  /** Blank-cell display behavior. */
  ExcelChartDisplayBlanksAs displayBlanksAs();

  /** Whether the chart should ignore hidden cells. */
  Boolean plotOnlyVisibleCells();

  /** Whether the chart should vary colors automatically. */
  Boolean varyColors();

  /** Requested chart series in display order. */
  List<Series> series();

  /** Supported simple bar chart input. */
  record Bar(
      String name,
      DrawingAnchorInput anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      Boolean varyColors,
      ExcelChartBarDirection barDirection,
      List<Series> series)
      implements ChartInput {
    public Bar {
      requireName(name);
      Objects.requireNonNull(anchor, "anchor must not be null");
      title = title == null ? new Title.None() : title;
      legend = legend == null ? new Legend.Visible(ExcelChartLegendPosition.RIGHT) : legend;
      displayBlanksAs = displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs;
      plotOnlyVisibleCells = plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells;
      varyColors = varyColors == null ? Boolean.FALSE : varyColors;
      barDirection = barDirection == null ? ExcelChartBarDirection.COLUMN : barDirection;
      series = copySeries(series);
    }
  }

  /** Supported simple line chart input. */
  record Line(
      String name,
      DrawingAnchorInput anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      Boolean varyColors,
      List<Series> series)
      implements ChartInput {
    public Line {
      requireName(name);
      Objects.requireNonNull(anchor, "anchor must not be null");
      title = title == null ? new Title.None() : title;
      legend = legend == null ? new Legend.Visible(ExcelChartLegendPosition.RIGHT) : legend;
      displayBlanksAs = displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs;
      plotOnlyVisibleCells = plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells;
      varyColors = varyColors == null ? Boolean.FALSE : varyColors;
      series = copySeries(series);
    }
  }

  /** Supported simple pie chart input. */
  record Pie(
      String name,
      DrawingAnchorInput anchor,
      Title title,
      Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      Boolean plotOnlyVisibleCells,
      Boolean varyColors,
      Integer firstSliceAngle,
      List<Series> series)
      implements ChartInput {
    public Pie {
      requireName(name);
      Objects.requireNonNull(anchor, "anchor must not be null");
      title = title == null ? new Title.None() : title;
      legend = legend == null ? new Legend.Visible(ExcelChartLegendPosition.RIGHT) : legend;
      displayBlanksAs = displayBlanksAs == null ? ExcelChartDisplayBlanksAs.GAP : displayBlanksAs;
      plotOnlyVisibleCells = plotOnlyVisibleCells == null ? Boolean.TRUE : plotOnlyVisibleCells;
      varyColors = varyColors == null ? Boolean.FALSE : varyColors;
      if (firstSliceAngle != null && (firstSliceAngle < 0 || firstSliceAngle > 360)) {
        throw new IllegalArgumentException("firstSliceAngle must be between 0 and 360");
      }
      series = copySeries(series);
    }
  }

  /** Requested chart title definition. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Title.None.class, name = "NONE"),
    @JsonSubTypes.Type(value = Title.Text.class, name = "TEXT"),
    @JsonSubTypes.Type(value = Title.Formula.class, name = "FORMULA")
  })
  sealed interface Title permits Title.None, Title.Text, Title.Formula {

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
  sealed interface Legend permits Legend.Hidden, Legend.Visible {

    /** Hide the legend entirely. */
    record Hidden() implements Legend {}

    /** Show the legend at one position. */
    record Visible(ExcelChartLegendPosition position) implements Legend {
      public Visible {
        Objects.requireNonNull(position, "position must not be null");
      }
    }
  }

  /** One authored chart series with workbook-bound sources. */
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
