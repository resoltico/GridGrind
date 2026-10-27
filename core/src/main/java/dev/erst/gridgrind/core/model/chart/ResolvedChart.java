package dev.erst.gridgrind.core.model.chart;

import static dev.erst.gridgrind.core.model.chart.ChartSpec.copyNonEmptyValues;
import static dev.erst.gridgrind.core.vocabulary.TextValidation.requireNonBlank;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Materialized chart whose title sources have become text; source-free options remain shared. */
public record ResolvedChart(
    String name,
    DrawingAnchorSpec.TwoCell anchor,
    ResolvedChartTitle title,
    ChartLegendSpec legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    List<Plot> plots) {
  public ResolvedChart {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    plots = copyNonEmptyValues(plots, "plots");
  }

  /** One authored chart series. */
  public record Series(
      ResolvedChartTitle title,
      ChartDataSourceSpec categories,
      ChartDataSourceSpec values,
      Optional<Boolean> smooth,
      Optional<ExcelChartMarkerStyle> markerStyle,
      Optional<Short> markerSize,
      Optional<Long> explosion) {
    public Series {
      title = title == null ? new ChartTitleSpec.None() : title;
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

    /** Convenience overload for authored series that only need title, categories, and values. */
    public Series(
        ResolvedChartTitle title, ChartDataSourceSpec categories, ChartDataSourceSpec values) {
      this(
          title,
          categories,
          values,
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty());
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
      boolean varyColors,
      ExcelChartGrouping grouping,
      List<ChartAxisSpec> axes,
      List<Series> series)
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
      Optional<Integer> gapDepth,
      List<ChartAxisSpec> axes,
      List<Series> series)
      implements Plot {
    public Area3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Bar(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapWidth,
      Optional<Integer> overlap,
      List<ChartAxisSpec> axes,
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

  public record Bar3D(
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      ExcelChartBarGrouping grouping,
      Optional<Integer> gapDepth,
      Optional<Integer> gapWidth,
      Optional<ExcelChartBarShape> shape,
      List<ChartAxisSpec> axes,
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

  public record Doughnut(
      boolean varyColors,
      Optional<Integer> firstSliceAngle,
      Optional<Integer> holeSize,
      List<Series> series)
      implements Plot {
    public Doughnut {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      Objects.requireNonNull(holeSize, "holeSize must not be null");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Line(
      boolean varyColors,
      ExcelChartGrouping grouping,
      List<ChartAxisSpec> axes,
      List<Series> series)
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
      Optional<Integer> gapDepth,
      List<ChartAxisSpec> axes,
      List<Series> series)
      implements Plot {
    public Line3D {
      Objects.requireNonNull(grouping, "grouping must not be null");
      Objects.requireNonNull(gapDepth, "gapDepth must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Pie(boolean varyColors, Optional<Integer> firstSliceAngle, List<Series> series)
      implements Plot {
    public Pie {
      Objects.requireNonNull(firstSliceAngle, "firstSliceAngle must not be null");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Pie3D(boolean varyColors, List<Series> series) implements Plot {
    public Pie3D {
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Radar(
      boolean varyColors, ExcelChartRadarStyle style, List<ChartAxisSpec> axes, List<Series> series)
      implements Plot {
    public Radar {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Scatter(
      boolean varyColors,
      ExcelChartScatterStyle style,
      List<ChartAxisSpec> axes,
      List<Series> series)
      implements Plot {
    public Scatter {
      Objects.requireNonNull(style, "style must not be null");
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Surface(
      boolean varyColors, boolean wireframe, List<ChartAxisSpec> axes, List<Series> series)
      implements Plot {
    public Surface {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }

  public record Surface3D(
      boolean varyColors, boolean wireframe, List<ChartAxisSpec> axes, List<Series> series)
      implements Plot {
    public Surface3D {
      axes = copyNonEmptyValues(axes, "axes");
      series = copyNonEmptyValues(series, "series");
    }
  }
}
