package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ResolvedInspectionStep;
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
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
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
    DrawingAnchorSpec.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 6, 10);
    ResolvedChart.Series definitionSeries =
        new ResolvedChart.Series(
            new ChartTitleSpec.None(),
            ExcelChartTestSupport.ref("A2:A4"),
            ExcelChartTestSupport.ref("B2:B4"),
            Optional.empty(),
            Optional.of(ExcelChartMarkerStyle.CIRCLE),
            Optional.of((short) 8),
            Optional.of(45L));

    ResolvedChart pieDefinition =
        ExcelChartTestSupport.pieChart(
            "OpsPie",
            anchor,
            new ResolvedChartTitle.Text("Share"),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            45,
            List.of(definitionSeries));
    ResolvedChart.Pie piePlot =
        assertInstanceOf(ResolvedChart.Pie.class, pieDefinition.plots().getFirst());
    assertEquals(Optional.of(45), piePlot.firstSliceAngle());

    ResolvedChart noAnglePie =
        ExcelChartTestSupport.pieChart(
            "NullAnglePie",
            anchor,
            new ChartTitleSpec.None(),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            false,
            null,
            List.of(definitionSeries));
    assertEquals(
        Optional.empty(),
        assertInstanceOf(ResolvedChart.Pie.class, noAnglePie.plots().getFirst()).firstSliceAngle());

    assertThrows(IllegalArgumentException.class, () -> new ResolvedChartTitle.Text(" "));
    assertThrows(IllegalArgumentException.class, () -> ExcelChartTestSupport.ref(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelChartTestSupport.lineChart(
                "OpsLine",
                anchor,
                new ChartTitleSpec.None(),
                new ChartLegendSpec.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                List.of()));

    ResolvedInspectionStep getCharts =
        new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
            new StepId("charts"),
            new ChartSelector.AllOnSheet(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Charts")),
            new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetCharts());
    assertEquals("charts", getCharts.stepId().value());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                new StepId(" "),
                new ChartSelector.AllOnSheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Charts")),
                new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetCharts()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                new StepId("charts"),
                new ChartSelector.AllOnSheet(new dev.erst.gridgrind.core.vocabulary.SheetName(" ")),
                new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery.GetCharts()));
  }

  @Test
  void snapshotShapesAreTopLevelChartsWithTypedPlots() {
    DrawingAnchorSpec.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 6, 10);
    ChartReport.Axis categoryAxis =
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true);
    ChartReport.Axis valueAxis =
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.MIN,
            false);
    ChartReport.Series lineSeries =
        new ChartReport.Series(
            new ChartReport.Title.Formula("Charts!$B$1", "Plan"),
            new ChartReport.DataSource.StringReference("Charts!$A$2:$A$4", List.of("Jan", "Feb")),
            new ChartReport.DataSource.NumericReference(
                "Charts!$B$2:$B$4", Optional.of("0.0"), List.of("10", "18")),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.CIRCLE),
            Optional.of((short) 8),
            Optional.empty());
    ChartReport lineSnapshot =
        new ChartReport(
            "OpsLine",
            DrawingGeometryFixtures.report(anchor),
            new ChartReport.Title.None(),
            new ChartReport.Legend.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            true,
            List.of(
                new ChartReport.Line(
                    false,
                    ExcelChartGrouping.STANDARD,
                    List.of(categoryAxis, valueAxis),
                    List.of(lineSeries))));
    ChartReport.Plot linePlot = lineSnapshot.plots().getFirst();
    assertInstanceOf(ChartReport.Line.class, linePlot);

    ChartReport unsupported =
        new ChartReport(
            "OpsUnsupported",
            DrawingGeometryFixtures.report(anchor),
            new ChartReport.Title.Text("Broken"),
            new ChartReport.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            false,
            List.of(new ChartReport.Unsupported("AREA", "unsupported")));
    assertEquals(
        "AREA",
        assertInstanceOf(ChartReport.Unsupported.class, unsupported.plots().getFirst())
            .plotTypeToken());

    List<ChartReport> chartSnapshots = new ArrayList<>(List.of(lineSnapshot, unsupported));
    dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult chartsResult =
        new dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult(
            new StepId("charts"),
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Charts"),
            chartSnapshots);
    chartSnapshots.clear();
    assertEquals(2, chartsResult.charts().size());

    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.DataSource.NumericReference("B2:B4", Optional.of(" "), List.of("1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartReport.DataSource.NumericLiteral(Optional.of(" "), List.of("1")));
    assertThrows(
        NullPointerException.class, () -> new ChartReport.Title.Formula("Charts!$B$1", null));
  }

  @Test
  void preparedModelsAndAllModeledPlotFamiliesRetainCanonicalState() throws Exception {
    DrawingAnchorSpec.TwoCell anchor = ExcelChartTestSupport.anchor(2, 2, 8, 14);
    List<ChartAxisSpec> categoryAxes =
        List.of(
            new ChartAxisSpec(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartAxisSpec(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.MIN,
                false));
    List<ChartAxisSpec> dateAxes =
        List.of(
            new ChartAxisSpec(
                ExcelChartAxisKind.DATE,
                ExcelChartAxisPosition.TOP,
                ExcelChartAxisCrosses.MAX,
                true),
            new ChartAxisSpec(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.RIGHT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true));
    List<ChartAxisSpec> scatterAxes =
        List.of(
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
    List<ChartAxisSpec> surfaceAxes =
        List.of(
            new ChartAxisSpec(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartAxisSpec(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.MIN,
                true),
            new ChartAxisSpec(
                ExcelChartAxisKind.SERIES,
                ExcelChartAxisPosition.RIGHT,
                ExcelChartAxisCrosses.MAX,
                false));
    ResolvedChart.Series definitionSeries =
        new ResolvedChart.Series(
            new ResolvedChartTitle.Text("Plan"),
            ExcelChartTestSupport.ref("A2:A4"),
            ExcelChartTestSupport.ref("B2:B4"),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.TRIANGLE),
            Optional.of((short) 10),
            Optional.of(25L));

    List<ResolvedChart.Plot> definitionPlots =
        List.of(
            new ResolvedChart.Area(
                true, ExcelChartGrouping.STACKED, categoryAxes, List.of(definitionSeries)),
            new ResolvedChart.Area3D(
                false,
                ExcelChartGrouping.PERCENT_STACKED,
                Optional.of(120),
                dateAxes,
                List.of(definitionSeries)),
            new ResolvedChart.Bar(
                true,
                ExcelChartBarDirection.BAR,
                ExcelChartBarGrouping.STACKED,
                Optional.of(90),
                Optional.of(40),
                categoryAxes,
                List.of(definitionSeries)),
            new ResolvedChart.Bar3D(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.PERCENT_STACKED,
                Optional.of(110),
                Optional.of(85),
                Optional.of(ExcelChartBarShape.CYLINDER),
                categoryAxes,
                List.of(definitionSeries)),
            new ResolvedChart.Doughnut(
                true, Optional.of(30), Optional.of(65), List.of(definitionSeries)),
            new ResolvedChart.Line(
                false, ExcelChartGrouping.STANDARD, categoryAxes, List.of(definitionSeries)),
            new ResolvedChart.Line3D(
                true,
                ExcelChartGrouping.STACKED,
                Optional.of(175),
                categoryAxes,
                List.of(definitionSeries)),
            new ResolvedChart.Pie(true, Optional.of(120), List.of(definitionSeries)),
            new ResolvedChart.Pie3D(false, List.of(definitionSeries)),
            new ResolvedChart.Radar(
                true, ExcelChartRadarStyle.MARKER, categoryAxes, List.of(definitionSeries)),
            new ResolvedChart.Scatter(
                false,
                ExcelChartScatterStyle.SMOOTH_MARKER,
                scatterAxes,
                List.of(definitionSeries)),
            new ResolvedChart.Surface(true, true, surfaceAxes, List.of(definitionSeries)),
            new ResolvedChart.Surface3D(false, false, surfaceAxes, List.of(definitionSeries)));
    assertEquals(13, definitionPlots.size());
    assertInstanceOf(ResolvedChart.Surface3D.class, definitionPlots.getLast());

    ChartReport.Series snapshotSeries =
        new ChartReport.Series(
            new ChartReport.Title.Text("Plan"),
            new ChartReport.DataSource.StringReference(
                "Charts!$A$2:$A$4", List.of("Jan", "Feb", "Mar")),
            new ChartReport.DataSource.NumericReference(
                "Charts!$B$2:$B$4", Optional.of("0.0"), List.of("10", "18", "15")),
            Optional.of(true),
            Optional.of(ExcelChartMarkerStyle.STAR),
            Optional.of((short) 11),
            Optional.of(50L));
    List<ChartReport.Axis> snapshotCategoryAxes =
        List.of(
            new ChartReport.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartReport.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.MIN,
                false));
    List<ChartReport.Axis> snapshotSurfaceAxes =
        List.of(
            new ChartReport.Axis(
                ExcelChartAxisKind.CATEGORY,
                ExcelChartAxisPosition.BOTTOM,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartReport.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartReport.Axis(
                ExcelChartAxisKind.SERIES,
                ExcelChartAxisPosition.RIGHT,
                ExcelChartAxisCrosses.MAX,
                false));
    List<ChartReport.Plot> snapshotPlots =
        List.of(
            new ChartReport.Area(
                true, ExcelChartGrouping.STANDARD, snapshotCategoryAxes, List.of(snapshotSeries)),
            new ChartReport.Area3D(
                false,
                ExcelChartGrouping.STACKED,
                Optional.of(110),
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ChartReport.Bar(
                true,
                ExcelChartBarDirection.BAR,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.of(75),
                Optional.of(15),
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ChartReport.Bar3D(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.PERCENT_STACKED,
                Optional.of(130),
                Optional.of(80),
                Optional.of(ExcelChartBarShape.PYRAMID),
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ChartReport.Doughnut(
                true, Optional.of(45), Optional.of(55), List.of(snapshotSeries)),
            new ChartReport.Line(
                false,
                ExcelChartGrouping.PERCENT_STACKED,
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ChartReport.Line3D(
                true,
                ExcelChartGrouping.STACKED,
                Optional.of(150),
                snapshotCategoryAxes,
                List.of(snapshotSeries)),
            new ChartReport.Pie(true, Optional.of(90), List.of(snapshotSeries)),
            new ChartReport.Pie3D(false, List.of(snapshotSeries)),
            new ChartReport.Radar(
                true, ExcelChartRadarStyle.FILLED, snapshotCategoryAxes, List.of(snapshotSeries)),
            new ChartReport.Scatter(
                false,
                ExcelChartScatterStyle.LINE_MARKER,
                List.of(
                    new ChartReport.Axis(
                        ExcelChartAxisKind.VALUE,
                        ExcelChartAxisPosition.BOTTOM,
                        ExcelChartAxisCrosses.MIN,
                        true),
                    new ChartReport.Axis(
                        ExcelChartAxisKind.VALUE,
                        ExcelChartAxisPosition.LEFT,
                        ExcelChartAxisCrosses.AUTO_ZERO,
                        true)),
                List.of(snapshotSeries)),
            new ChartReport.Surface(true, true, snapshotSurfaceAxes, List.of(snapshotSeries)),
            new ChartReport.Surface3D(false, false, snapshotSurfaceAxes, List.of(snapshotSeries)),
            new ChartReport.Unsupported("CUSTOM", "detail"));
    assertEquals(14, snapshotPlots.size());
    assertInstanceOf(ChartReport.Unsupported.class, snapshotPlots.getLast());

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
              new ChartLegendSpec.Hidden(),
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
              new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
              ExcelChartDisplayBlanksAs.SPAN,
              false,
              true,
              mutableSeries);
      PreparedPieChart preparedPie =
          new PreparedPieChart(
              "OpsPie",
              anchor,
              new PreparedChartTitleNone(),
              new ChartLegendSpec.Hidden(),
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
