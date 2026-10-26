package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
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
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** End-to-end coverage for every modeled chart family in the engine surface. */
class ExcelChartFamilyCoverageTest {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void everyModeledChartFamilyRoundTripsThroughMutationSnapshotAndDrawingInventory()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);

      sheet
          .drawings()
          .setChart(
              chart(
                  "AreaOps",
                  ExcelChartTestSupport.anchor(1, 1, 7, 11),
                  new ResolvedChart.Area(
                      true,
                      ExcelChartGrouping.STACKED,
                      categoryAxes(),
                      List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "Area3DOps",
                  ExcelChartTestSupport.anchor(8, 1, 14, 11),
                  new ResolvedChart.Area3D(
                      false,
                      ExcelChartGrouping.PERCENT_STACKED,
                      Optional.of(140),
                      dateAxes(),
                      List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "BarOps",
                  ExcelChartTestSupport.anchor(15, 1, 21, 11),
                  new ResolvedChart.Bar(
                      true,
                      ExcelChartBarDirection.BAR,
                      ExcelChartBarGrouping.STACKED,
                      Optional.of(90),
                      Optional.of(40),
                      categoryAxes(),
                      List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "Bar3DOps",
                  ExcelChartTestSupport.anchor(22, 1, 28, 11),
                  new ResolvedChart.Bar3D(
                      false,
                      ExcelChartBarDirection.COLUMN,
                      ExcelChartBarGrouping.PERCENT_STACKED,
                      Optional.of(120),
                      Optional.of(80),
                      Optional.of(ExcelChartBarShape.CYLINDER),
                      categoryAxes(),
                      List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "DoughnutOps",
                  ExcelChartTestSupport.anchor(29, 1, 35, 11),
                  new ResolvedChart.Doughnut(
                      true,
                      Optional.of(35),
                      Optional.of(65),
                      List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, 30L)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "LineOps",
                  ExcelChartTestSupport.anchor(36, 1, 42, 11),
                  new ResolvedChart.Line(
                      false,
                      ExcelChartGrouping.PERCENT_STACKED,
                      categoryAxes(),
                      List.of(
                          series(
                              "Plan",
                              "A2:A4",
                              "B2:B4",
                              true,
                              ExcelChartMarkerStyle.CIRCLE,
                              (short) 8,
                              null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "Line3DOps",
                  ExcelChartTestSupport.anchor(43, 1, 49, 11),
                  new ResolvedChart.Line3D(
                      true,
                      ExcelChartGrouping.STACKED,
                      Optional.of(175),
                      categoryAxes(),
                      List.of(
                          series(
                              "Actual",
                              "A2:A4",
                              "C2:C4",
                              false,
                              ExcelChartMarkerStyle.DIAMOND,
                              (short) 9,
                              null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "PieOps",
                  ExcelChartTestSupport.anchor(50, 1, 56, 11),
                  new ResolvedChart.Pie(
                      true,
                      Optional.of(120),
                      List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, 45L)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "Pie3DOps",
                  ExcelChartTestSupport.anchor(57, 1, 63, 11),
                  new ResolvedChart.Pie3D(
                      false, List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, 10L)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "RadarOps",
                  ExcelChartTestSupport.anchor(64, 1, 70, 11),
                  new ResolvedChart.Radar(
                      true,
                      ExcelChartRadarStyle.MARKER,
                      categoryAxes(),
                      List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "ScatterOps",
                  ExcelChartTestSupport.anchor(71, 1, 77, 11),
                  new ResolvedChart.Scatter(
                      false,
                      ExcelChartScatterStyle.SMOOTH_MARKER,
                      scatterAxes(),
                      List.of(
                          series(
                              "Trend",
                              "B2:B4",
                              "C2:C4",
                              true,
                              ExcelChartMarkerStyle.STAR,
                              (short) 11,
                              null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "SurfaceOps",
                  ExcelChartTestSupport.anchor(78, 1, 84, 11),
                  new ResolvedChart.Surface(
                      true,
                      true,
                      surfaceAxes(),
                      List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet
          .drawings()
          .setChart(
              chart(
                  "Surface3DOps",
                  ExcelChartTestSupport.anchor(85, 1, 91, 11),
                  new ResolvedChart.Surface3D(
                      false,
                      false,
                      surfaceAxes(),
                      List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, null)))));

      List<ExcelChartSnapshot> charts = sheet.drawings().charts();
      assertEquals(13, charts.size());

      ExcelChartSnapshot.Area area =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "AreaOps"), ExcelChartSnapshot.Area.class);
      assertTrue(area.varyColors());
      assertEquals(ExcelChartGrouping.STACKED, area.grouping());

      ExcelChartSnapshot.Area3D area3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Area3DOps"), ExcelChartSnapshot.Area3D.class);
      assertEquals(Optional.of(140), area3D.gapDepth());
      assertEquals(ExcelChartAxisKind.DATE, area3D.axes().getFirst().kind());

      ExcelChartSnapshot.Bar bar =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "BarOps"), ExcelChartSnapshot.Bar.class);
      assertEquals(ExcelChartBarDirection.BAR, bar.barDirection());
      assertEquals(ExcelChartBarGrouping.STACKED, bar.grouping());
      assertEquals(Optional.of(40), bar.overlap());

      ExcelChartSnapshot.Bar3D bar3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Bar3DOps"), ExcelChartSnapshot.Bar3D.class);
      assertEquals(Optional.of(ExcelChartBarShape.CYLINDER), bar3D.shape());
      assertEquals(Optional.of(120), bar3D.gapDepth());

      ExcelChartSnapshot.Doughnut doughnut =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "DoughnutOps"),
              ExcelChartSnapshot.Doughnut.class);
      assertEquals(Optional.of(35), doughnut.firstSliceAngle());
      assertEquals(Optional.of(65), doughnut.holeSize());
      assertEquals(Optional.of(30L), doughnut.series().getFirst().explosion());

      ExcelChartSnapshot.Line line =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "LineOps"), ExcelChartSnapshot.Line.class);
      assertEquals(
          Optional.of(ExcelChartMarkerStyle.CIRCLE), line.series().getFirst().markerStyle());
      assertEquals(Optional.of(true), line.series().getFirst().smooth());

      ExcelChartSnapshot.Line3D line3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Line3DOps"), ExcelChartSnapshot.Line3D.class);
      assertEquals(Optional.of(175), line3D.gapDepth());
      assertEquals(Optional.of(false), line3D.series().getFirst().smooth());

      ExcelChartSnapshot.Pie pie =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "PieOps"), ExcelChartSnapshot.Pie.class);
      assertEquals(Optional.of(120), pie.firstSliceAngle());
      assertEquals(Optional.of(45L), pie.series().getFirst().explosion());

      ExcelChartSnapshot.Pie3D pie3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Pie3DOps"), ExcelChartSnapshot.Pie3D.class);
      assertFalse(pie3D.varyColors());
      assertEquals(Optional.of(10L), pie3D.series().getFirst().explosion());

      ExcelChartSnapshot.Radar radar =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "RadarOps"), ExcelChartSnapshot.Radar.class);
      assertEquals(ExcelChartRadarStyle.MARKER, radar.style());

      ExcelChartSnapshot.Scatter scatter =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "ScatterOps"), ExcelChartSnapshot.Scatter.class);
      assertEquals(ExcelChartScatterStyle.SMOOTH_MARKER, scatter.style());
      assertEquals(
          Optional.of(ExcelChartMarkerStyle.STAR), scatter.series().getFirst().markerStyle());
      assertTrue(scatter.axes().stream().allMatch(axis -> axis.kind() == ExcelChartAxisKind.VALUE));

      ExcelChartSnapshot.Surface surface =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "SurfaceOps"), ExcelChartSnapshot.Surface.class);
      assertTrue(surface.wireframe());
      assertTrue(
          surface.axes().stream().anyMatch(axis -> axis.kind() == ExcelChartAxisKind.SERIES));

      ExcelChartSnapshot.Surface3D surface3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Surface3DOps"),
              ExcelChartSnapshot.Surface3D.class);
      assertFalse(surface3D.wireframe());

      List<ExcelDrawingObjectSnapshot.Chart> drawingCharts =
          sheet.drawings().drawingObjects().stream()
              .map(snapshot -> assertInstanceOf(ExcelDrawingObjectSnapshot.Chart.class, snapshot))
              .toList();
      assertEquals(13, drawingCharts.size());
      assertTrue(
          drawingCharts.stream()
              .allMatch(snapshot -> snapshot.supported() && snapshot.plotTypeTokens().size() == 1));
      assertTrue(
          drawingCharts.stream()
              .anyMatch(
                  snapshot ->
                      "ScatterOps".equals(snapshot.name())
                          && snapshot.plotTypeTokens().equals(List.of("SCATTER"))));
    }
  }

  private static ResolvedChart chart(
      String name, DrawingAnchorSpec.TwoCell anchor, ResolvedChart.Plot plot) {
    return new ResolvedChart(
        name,
        anchor,
        new ResolvedChartTitle.Text(name),
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(plot));
  }

  private static ResolvedChart.Series series(
      String title,
      String categories,
      String values,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      Long explosion) {
    return new ResolvedChart.Series(
        new ResolvedChartTitle.Text(title),
        ExcelChartTestSupport.ref(categories),
        ExcelChartTestSupport.ref(values),
        Optional.ofNullable(smooth),
        Optional.ofNullable(markerStyle),
        Optional.ofNullable(markerSize),
        Optional.ofNullable(explosion));
  }

  private static List<ChartAxisSpec> categoryAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.MIN,
            true));
  }

  private static List<ChartAxisSpec> dateAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.DATE, ExcelChartAxisPosition.TOP, ExcelChartAxisCrosses.MAX, true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisSpec> scatterAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.MIN,
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
            ExcelChartAxisKind.VALUE, ExcelChartAxisPosition.LEFT, ExcelChartAxisCrosses.MIN, true),
        new ChartAxisSpec(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.MAX,
            true));
  }
}
