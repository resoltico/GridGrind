package dev.erst.gridgrind.core.fixture;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.util.List;
import java.util.Optional;

/** Shared authored inputs for extended chart preparation and source-binding cases. */
public final class ChartInputFixtures {
  private ChartInputFixtures() {}

  /** Provides the shared two-cell chart placement. */
  public static DrawingAnchorSpec.TwoCell anchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(1, 2, 0, 0),
        new DrawingMarkerSpec(8, 14, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  /** Supplies literal categories and values with the selected title source. */
  public static ChartSeriesSpec series(TextSourceSpec title) {
    return new ChartSeriesSpec(
        new ChartTitleSpec.Text(title),
        new ChartDataSourceSpec.StringLiteral(List.of("Jan", "Feb")),
        new ChartDataSourceSpec.NumericLiteral(List.of(10.0d, 18.0d)),
        Optional.of(true),
        Optional.of(ExcelChartMarkerStyle.DIAMOND),
        Optional.of((short) 6),
        Optional.of(4L));
  }

  /** Supplies all thirteen authored chart kinds in stable fixture order. */
  public static List<ChartPlotSpec> plots(ChartSeriesSpec series) {
    return List.of(
        new ChartPlotSpec.Area(false, ExcelChartGrouping.STANDARD, axes(), List.of(series)),
        new ChartPlotSpec.Area3D(
            false, ExcelChartGrouping.PERCENT_STACKED, Optional.of(24), axes(), List.of(series)),
        new ChartPlotSpec.Bar(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.of(60),
            Optional.of(0),
            axes(),
            List.of(series)),
        new ChartPlotSpec.Bar3D(
            true,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.STACKED,
            Optional.of(32),
            Optional.of(88),
            Optional.of(ExcelChartBarShape.CONE),
            axes(),
            List.of(series)),
        new ChartPlotSpec.Doughnut(true, Optional.of(30), Optional.of(55), List.of(series)),
        new ChartPlotSpec.Line(false, ExcelChartGrouping.STANDARD, axes(), List.of(series)),
        new ChartPlotSpec.Line3D(
            false, ExcelChartGrouping.STANDARD, Optional.of(18), axes(), List.of(series)),
        new ChartPlotSpec.Pie(true, Optional.of(90), List.of(series)),
        new ChartPlotSpec.Pie3D(true, List.of(series)),
        new ChartPlotSpec.Radar(false, ExcelChartRadarStyle.FILLED, axes(), List.of(series)),
        new ChartPlotSpec.Scatter(
            false, ExcelChartScatterStyle.SMOOTH_MARKER, scatterAxes(), List.of(series)),
        new ChartPlotSpec.Surface(false, true, surfaceAxes(), List.of(series)),
        new ChartPlotSpec.Surface3D(true, false, surfaceAxes(), List.of(series)));
  }

  private static List<ChartAxisSpec> axes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisSpec> scatterAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisSpec> surfaceAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }
}
