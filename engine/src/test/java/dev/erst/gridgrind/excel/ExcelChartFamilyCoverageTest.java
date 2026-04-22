package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end coverage for every modeled chart family in the engine surface. */
class ExcelChartFamilyCoverageTest {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void everyModeledChartFamilyRoundTripsThroughMutationSnapshotAndDrawingInventory()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);

      sheet.setChart(
          chart(
              "AreaOps",
              ExcelChartTestSupport.anchor(1, 1, 7, 11),
              new ExcelChartDefinition.Area(
                  true,
                  ExcelChartGrouping.STACKED,
                  categoryAxes(),
                  List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet.setChart(
          chart(
              "Area3DOps",
              ExcelChartTestSupport.anchor(8, 1, 14, 11),
              new ExcelChartDefinition.Area3D(
                  false,
                  ExcelChartGrouping.PERCENT_STACKED,
                  140,
                  dateAxes(),
                  List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, null)))));
      sheet.setChart(
          chart(
              "BarOps",
              ExcelChartTestSupport.anchor(15, 1, 21, 11),
              new ExcelChartDefinition.Bar(
                  true,
                  ExcelChartBarDirection.BAR,
                  ExcelChartBarGrouping.STACKED,
                  90,
                  40,
                  categoryAxes(),
                  List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet.setChart(
          chart(
              "Bar3DOps",
              ExcelChartTestSupport.anchor(22, 1, 28, 11),
              new ExcelChartDefinition.Bar3D(
                  false,
                  ExcelChartBarDirection.COLUMN,
                  ExcelChartBarGrouping.PERCENT_STACKED,
                  120,
                  80,
                  ExcelChartBarShape.CYLINDER,
                  categoryAxes(),
                  List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, null)))));
      sheet.setChart(
          chart(
              "DoughnutOps",
              ExcelChartTestSupport.anchor(29, 1, 35, 11),
              new ExcelChartDefinition.Doughnut(
                  true,
                  35,
                  65,
                  List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, 30L)))));
      sheet.setChart(
          chart(
              "LineOps",
              ExcelChartTestSupport.anchor(36, 1, 42, 11),
              new ExcelChartDefinition.Line(
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
      sheet.setChart(
          chart(
              "Line3DOps",
              ExcelChartTestSupport.anchor(43, 1, 49, 11),
              new ExcelChartDefinition.Line3D(
                  true,
                  ExcelChartGrouping.STACKED,
                  175,
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
      sheet.setChart(
          chart(
              "PieOps",
              ExcelChartTestSupport.anchor(50, 1, 56, 11),
              new ExcelChartDefinition.Pie(
                  true, 120, List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, 45L)))));
      sheet.setChart(
          chart(
              "Pie3DOps",
              ExcelChartTestSupport.anchor(57, 1, 63, 11),
              new ExcelChartDefinition.Pie3D(
                  false, List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, 10L)))));
      sheet.setChart(
          chart(
              "RadarOps",
              ExcelChartTestSupport.anchor(64, 1, 70, 11),
              new ExcelChartDefinition.Radar(
                  true,
                  ExcelChartRadarStyle.MARKER,
                  categoryAxes(),
                  List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet.setChart(
          chart(
              "ScatterOps",
              ExcelChartTestSupport.anchor(71, 1, 77, 11),
              new ExcelChartDefinition.Scatter(
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
      sheet.setChart(
          chart(
              "SurfaceOps",
              ExcelChartTestSupport.anchor(78, 1, 84, 11),
              new ExcelChartDefinition.Surface(
                  true,
                  true,
                  surfaceAxes(),
                  List.of(series("Plan", "A2:A4", "B2:B4", null, null, null, null)))));
      sheet.setChart(
          chart(
              "Surface3DOps",
              ExcelChartTestSupport.anchor(85, 1, 91, 11),
              new ExcelChartDefinition.Surface3D(
                  false,
                  false,
                  surfaceAxes(),
                  List.of(series("Actual", "A2:A4", "C2:C4", null, null, null, null)))));

      List<ExcelChartSnapshot> charts = sheet.charts();
      assertEquals(13, charts.size());

      ExcelChartSnapshot.Area area =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "AreaOps"), ExcelChartSnapshot.Area.class);
      assertTrue(area.varyColors());
      assertEquals(ExcelChartGrouping.STACKED, area.grouping());

      ExcelChartSnapshot.Area3D area3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Area3DOps"), ExcelChartSnapshot.Area3D.class);
      assertEquals(140, area3D.gapDepth());
      assertEquals(ExcelChartAxisKind.DATE, area3D.axes().getFirst().kind());

      ExcelChartSnapshot.Bar bar =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "BarOps"), ExcelChartSnapshot.Bar.class);
      assertEquals(ExcelChartBarDirection.BAR, bar.barDirection());
      assertEquals(ExcelChartBarGrouping.STACKED, bar.grouping());
      assertEquals(40, bar.overlap());

      ExcelChartSnapshot.Bar3D bar3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Bar3DOps"), ExcelChartSnapshot.Bar3D.class);
      assertEquals(ExcelChartBarShape.CYLINDER, bar3D.shape());
      assertEquals(120, bar3D.gapDepth());

      ExcelChartSnapshot.Doughnut doughnut =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "DoughnutOps"),
              ExcelChartSnapshot.Doughnut.class);
      assertEquals(35, doughnut.firstSliceAngle());
      assertEquals(65, doughnut.holeSize());
      assertEquals(30L, doughnut.series().getFirst().explosion());

      ExcelChartSnapshot.Line line =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "LineOps"), ExcelChartSnapshot.Line.class);
      assertEquals(ExcelChartMarkerStyle.CIRCLE, line.series().getFirst().markerStyle());
      assertEquals(true, line.series().getFirst().smooth());

      ExcelChartSnapshot.Line3D line3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Line3DOps"), ExcelChartSnapshot.Line3D.class);
      assertEquals(175, line3D.gapDepth());
      assertEquals(false, line3D.series().getFirst().smooth());

      ExcelChartSnapshot.Pie pie =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "PieOps"), ExcelChartSnapshot.Pie.class);
      assertEquals(120, pie.firstSliceAngle());
      assertEquals(45L, pie.series().getFirst().explosion());

      ExcelChartSnapshot.Pie3D pie3D =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "Pie3DOps"), ExcelChartSnapshot.Pie3D.class);
      assertFalse(pie3D.varyColors());
      assertEquals(10L, pie3D.series().getFirst().explosion());

      ExcelChartSnapshot.Radar radar =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "RadarOps"), ExcelChartSnapshot.Radar.class);
      assertEquals(ExcelChartRadarStyle.MARKER, radar.style());

      ExcelChartSnapshot.Scatter scatter =
          ExcelChartTestSupport.singlePlot(
              ExcelChartTestSupport.chart(charts, "ScatterOps"), ExcelChartSnapshot.Scatter.class);
      assertEquals(ExcelChartScatterStyle.SMOOTH_MARKER, scatter.style());
      assertEquals(ExcelChartMarkerStyle.STAR, scatter.series().getFirst().markerStyle());
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
          sheet.drawingObjects().stream()
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

  private static ExcelChartDefinition chart(
      String name, ExcelDrawingAnchor.TwoCell anchor, ExcelChartDefinition.Plot plot) {
    return new ExcelChartDefinition(
        name,
        anchor,
        new ExcelChartDefinition.Title.Text(name),
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(plot));
  }

  private static ExcelChartDefinition.Series series(
      String title,
      String categories,
      String values,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      Long explosion) {
    return new ExcelChartDefinition.Series(
        new ExcelChartDefinition.Title.Text(title),
        ExcelChartTestSupport.ref(categories),
        ExcelChartTestSupport.ref(values),
        smooth,
        markerStyle,
        markerSize,
        explosion);
  }

  private static List<ExcelChartDefinition.Axis> categoryAxes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.MIN,
            true));
  }

  private static List<ExcelChartDefinition.Axis> dateAxes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.DATE, ExcelChartAxisPosition.TOP, ExcelChartAxisCrosses.MAX, true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> scatterAxes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.MIN,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> surfaceAxes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE, ExcelChartAxisPosition.LEFT, ExcelChartAxisCrosses.MIN, true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.MAX,
            true));
  }
}
