package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Value-object coverage for the rebuilt chart definition and snapshot model. */
class ExcelChartValueObjectTest {
  @Test
  void definitionShapesAndCommandWrappersValidateCanonicalState() {
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 6, 10);
    ExcelChartDefinition.Series definitionSeries =
        new ExcelChartDefinition.Series(
            new ExcelChartDefinition.Title.None(),
            ExcelChartTestSupport.ref("A2:A4"),
            ExcelChartTestSupport.ref("B2:B4"),
            null,
            ExcelChartMarkerStyle.CIRCLE,
            (short) 8,
            45L);

    ExcelChartDefinition pieDefinition =
        ExcelChartTestSupport.pieChart(
            "OpsPie",
            anchor,
            new ExcelChartDefinition.Title.Text("Share"),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            45,
            List.of(definitionSeries));
    ExcelChartDefinition.Pie piePlot =
        assertInstanceOf(ExcelChartDefinition.Pie.class, pieDefinition.plots().getFirst());
    assertEquals(45, piePlot.firstSliceAngle());

    ExcelChartDefinition noAnglePie =
        ExcelChartTestSupport.pieChart(
            "NullAnglePie",
            anchor,
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            false,
            null,
            List.of(definitionSeries));
    assertNull(
        assertInstanceOf(ExcelChartDefinition.Pie.class, noAnglePie.plots().getFirst())
            .firstSliceAngle());

    assertThrows(IllegalArgumentException.class, () -> new ExcelChartDefinition.Title.Text(" "));
    assertThrows(IllegalArgumentException.class, () -> ExcelChartTestSupport.ref(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelChartTestSupport.lineChart(
                "OpsLine",
                anchor,
                new ExcelChartDefinition.Title.None(),
                new ExcelChartDefinition.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                List.of()));

    WorkbookCommand.SetChart setChart = new WorkbookCommand.SetChart("Charts", pieDefinition);
    assertEquals("Charts", setChart.sheetName());
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetChart(" ", pieDefinition));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetChart("Charts", null));

    WorkbookReadCommand.GetCharts getCharts = new WorkbookReadCommand.GetCharts("charts", "Charts");
    assertEquals("charts", getCharts.stepId());
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadCommand.GetCharts(" ", "Charts"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadCommand.GetCharts("charts", " "));
  }

  @Test
  void snapshotShapesAreTopLevelChartsWithTypedPlots() {
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 6, 10);
    ExcelChartSnapshot.Axis categoryAxis =
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true);
    ExcelChartSnapshot.Axis valueAxis =
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.MIN,
            false);
    ExcelChartSnapshot.Series lineSeries =
        new ExcelChartSnapshot.Series(
            new ExcelChartSnapshot.Title.Formula("Charts!$B$1", "Plan"),
            new ExcelChartSnapshot.DataSource.StringReference(
                "Charts!$A$2:$A$4", List.of("Jan", "Feb")),
            new ExcelChartSnapshot.DataSource.NumericReference(
                "Charts!$B$2:$B$4", "0.0", List.of("10", "18")),
            true,
            ExcelChartMarkerStyle.CIRCLE,
            (short) 8,
            null);
    ExcelChartSnapshot lineSnapshot =
        new ExcelChartSnapshot(
            "OpsLine",
            anchor,
            new ExcelChartSnapshot.Title.None(),
            new ExcelChartSnapshot.Legend.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            true,
            List.of(
                new ExcelChartSnapshot.Line(
                    false,
                    ExcelChartGrouping.STANDARD,
                    List.of(categoryAxis, valueAxis),
                    List.of(lineSeries))));
    ExcelChartSnapshot.Plot linePlot = lineSnapshot.plots().getFirst();
    assertInstanceOf(ExcelChartSnapshot.Line.class, linePlot);

    ExcelChartSnapshot unsupported =
        new ExcelChartSnapshot(
            "OpsUnsupported",
            anchor,
            new ExcelChartSnapshot.Title.Text("Broken"),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            false,
            List.of(new ExcelChartSnapshot.Unsupported("AREA", "unsupported")));
    assertEquals(
        "AREA",
        assertInstanceOf(ExcelChartSnapshot.Unsupported.class, unsupported.plots().getFirst())
            .plotTypeToken());

    List<ExcelChartSnapshot> chartSnapshots = new ArrayList<>(List.of(lineSnapshot, unsupported));
    WorkbookReadResult.ChartsResult chartsResult =
        new WorkbookReadResult.ChartsResult("charts", "Charts", chartSnapshots);
    chartSnapshots.clear();
    assertEquals(2, chartsResult.charts().size());

    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelChartSnapshot.DataSource.NumericReference("B2:B4", " ", List.of("1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelChartSnapshot.DataSource.NumericLiteral(" ", List.of("1")));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelChartSnapshot.Title.Formula("Charts!$B$1", null));
  }

  @Test
  void preparedModelsAndAllModeledPlotFamiliesRetainCanonicalState() throws Exception {
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(2, 2, 8, 14);
    List<ExcelChartDefinition.Axis> categoryAxes =
        List.of(
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.MIN,
                false));
    List<ExcelChartDefinition.Axis> dateAxes =
        List.of(
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.DATE,
                ExcelChartAxisPosition.TOP,
                ExcelChartAxisCrosses.MAX,
                true),
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.RIGHT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true));
    List<ExcelChartDefinition.Axis> scatterAxes =
        List.of(
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
    List<ExcelChartDefinition.Axis> surfaceAxes =
        List.of(
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.MIN,
                true),
            new ExcelChartDefinition.Axis(
                ExcelChartAxisKind.SERIES,
                ExcelChartAxisPosition.RIGHT,
                ExcelChartAxisCrosses.MAX,
                false));
    ExcelChartDefinition.Series definitionSeries =
        new ExcelChartDefinition.Series(
            new ExcelChartDefinition.Title.Text("Plan"),
            ExcelChartTestSupport.ref("A2:A4"),
            ExcelChartTestSupport.ref("B2:B4"),
            true,
            ExcelChartMarkerStyle.TRIANGLE,
            (short) 10,
            25L);

    List<ExcelChartDefinition.Plot> definitionPlots =
        List.of(
            new ExcelChartDefinition.Area(
                true, ExcelChartGrouping.STACKED, categoryAxes, List.of(definitionSeries)),
            new ExcelChartDefinition.Area3D(
                false,
                ExcelChartGrouping.PERCENT_STACKED,
                120,
                dateAxes,
                List.of(definitionSeries)),
            new ExcelChartDefinition.Bar(
                true,
                ExcelChartBarDirection.BAR,
                ExcelChartBarGrouping.STACKED,
                90,
                40,
                categoryAxes,
                List.of(definitionSeries)),
            new ExcelChartDefinition.Bar3D(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.PERCENT_STACKED,
                110,
                85,
                ExcelChartBarShape.CYLINDER,
                categoryAxes,
                List.of(definitionSeries)),
            new ExcelChartDefinition.Doughnut(true, 30, 65, List.of(definitionSeries)),
            new ExcelChartDefinition.Line(
                false, ExcelChartGrouping.STANDARD, categoryAxes, List.of(definitionSeries)),
            new ExcelChartDefinition.Line3D(
                true, ExcelChartGrouping.STACKED, 175, categoryAxes, List.of(definitionSeries)),
            new ExcelChartDefinition.Pie(true, 120, List.of(definitionSeries)),
            new ExcelChartDefinition.Pie3D(false, List.of(definitionSeries)),
            new ExcelChartDefinition.Radar(
                true, ExcelChartRadarStyle.MARKER, categoryAxes, List.of(definitionSeries)),
            new ExcelChartDefinition.Scatter(
                false,
                ExcelChartScatterStyle.SMOOTH_MARKER,
                scatterAxes,
                List.of(definitionSeries)),
            new ExcelChartDefinition.Surface(true, true, surfaceAxes, List.of(definitionSeries)),
            new ExcelChartDefinition.Surface3D(
                false, false, surfaceAxes, List.of(definitionSeries)));
    assertEquals(13, definitionPlots.size());
    assertInstanceOf(ExcelChartDefinition.Surface3D.class, definitionPlots.getLast());

    ExcelChartSnapshot.Series snapshotSeries =
        new ExcelChartSnapshot.Series(
            new ExcelChartSnapshot.Title.Text("Plan"),
            new ExcelChartSnapshot.DataSource.StringReference(
                "Charts!$A$2:$A$4", List.of("Jan", "Feb", "Mar")),
            new ExcelChartSnapshot.DataSource.NumericReference(
                "Charts!$B$2:$B$4", "0.0", List.of("10", "18", "15")),
            true,
            ExcelChartMarkerStyle.STAR,
            (short) 11,
            50L);
    List<ExcelChartSnapshot.Axis> snapshotCategoryAxes =
        List.of(
            new ExcelChartSnapshot.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ExcelChartSnapshot.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.MIN,
                false));
    List<ExcelChartSnapshot.Axis> snapshotSurfaceAxes =
        List.of(
            new ExcelChartSnapshot.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ExcelChartSnapshot.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ExcelChartSnapshot.Axis(
                ExcelChartAxisKind.SERIES,
                ExcelChartAxisPosition.RIGHT,
                ExcelChartAxisCrosses.MAX,
                false));
    List<ExcelChartSnapshot.Plot> snapshotPlots =
        List.of(
            new ExcelChartSnapshot.Area(
                true, ExcelChartGrouping.STANDARD, snapshotCategoryAxes, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Area3D(
                false,
                ExcelChartGrouping.STACKED,
                110,
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ExcelChartSnapshot.Bar(
                true,
                ExcelChartBarDirection.BAR,
                ExcelChartBarGrouping.CLUSTERED,
                75,
                15,
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ExcelChartSnapshot.Bar3D(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.PERCENT_STACKED,
                130,
                80,
                ExcelChartBarShape.PYRAMID,
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ExcelChartSnapshot.Doughnut(true, 45, 55, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Line(
                false,
                ExcelChartGrouping.PERCENT_STACKED,
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ExcelChartSnapshot.Line3D(
                true,
                ExcelChartGrouping.STACKED,
                150,
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ExcelChartSnapshot.Pie(true, 90, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Pie3D(false, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Radar(
                true, ExcelChartRadarStyle.FILLED, snapshotCategoryAxes, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Scatter(
                false,
                ExcelChartScatterStyle.LINE_MARKER,
                List.of(
                    new ExcelChartSnapshot.Axis(
                        ExcelChartAxisKind.VALUE,
                        ExcelChartAxisPosition.BOTTOM,
                        ExcelChartAxisCrosses.MIN,
                        true),
                    new ExcelChartSnapshot.Axis(
                        ExcelChartAxisKind.VALUE,
                        ExcelChartAxisPosition.LEFT,
                        ExcelChartAxisCrosses.AUTO_ZERO,
                        true)),
                List.of(snapshotSeries)),
            new ExcelChartSnapshot.Surface(
                true, true, snapshotSurfaceAxes, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Surface3D(
                false, false, snapshotSurfaceAxes, List.of(snapshotSeries)),
            new ExcelChartSnapshot.Unsupported("CUSTOM", "detail"));
    assertEquals(14, snapshotPlots.size());
    assertInstanceOf(ExcelChartSnapshot.Unsupported.class, snapshotPlots.getLast());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ResolvedAreaReference resolvedArea =
          new ResolvedAreaReference(
              sheet, new AreaReference("Charts!$A$2:$B$4", SpreadsheetVersion.EXCEL2007));
      ResolvedChartSource resolvedChartSource =
          new ResolvedChartSource(
              "Charts!$A$2:$B$4",
              sheet,
              resolvedArea.areaReference(),
              false,
              List.of("Jan", "Feb", "Mar"),
              List.of(10d, 18d, 15d));
      assertEquals("Charts!$A$2:$B$4", resolvedChartSource.referenceFormula());

      PreparedSeriesTitle preparedSeriesTitle =
          new PreparedSeriesTitleFormula("Plan", new CellReference("Charts", 0, 1, true, true));
      PreparedChartSeries preparedSeries =
          new PreparedChartSeries(
              preparedSeriesTitle,
              XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb", "Mar"}),
              XDDFDataSourcesFactory.fromArray(new Double[] {10d, 18d, 15d}));
      List<PreparedChartSeries> mutableSeries = new ArrayList<>(List.of(preparedSeries));
      PreparedBarChart preparedBar =
          new PreparedBarChart(
              "OpsBar",
              anchor,
              new PreparedChartTitleText("Roadmap"),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              false,
              ExcelChartBarDirection.COLUMN,
              mutableSeries);
      PreparedLineChart preparedLine =
          new PreparedLineChart(
              "OpsLine",
              anchor,
              new PreparedChartTitleFormula("1.0", new CellReference("Charts", 0, 3, true, true)),
              new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT),
              ExcelChartDisplayBlanksAs.SPAN,
              false,
              true,
              mutableSeries);
      PreparedPieChart preparedPie =
          new PreparedPieChart(
              "OpsPie",
              anchor,
              new PreparedChartTitleNone(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.ZERO,
              true,
              true,
              25,
              mutableSeries);
      mutableSeries.clear();
      assertEquals(1, preparedBar.series().size());
      assertEquals("1.0", ((PreparedChartTitleFormula) preparedLine.title()).cachedText());
      assertEquals(25, preparedPie.firstSliceAngle());
    }
  }
}
