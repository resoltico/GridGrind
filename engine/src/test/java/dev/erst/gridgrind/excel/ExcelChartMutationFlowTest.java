package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end mutation and introspection coverage for authored charts. */
class ExcelChartMutationFlowTest {
  @Test
  void lineAndPieChartsExerciseMutationExecutorAndIntrospection() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartTestSupport.seedChartNamedRanges(workbook, "Charts");

      executor.apply(
          workbook,
          new WorkbookCommand.SetChart(
              "Charts",
              lineChartDefinition(
                  "OpsLine",
                  ExcelChartTestSupport.anchor(4, 1, 10, 16),
                  new ExcelChartDefinition.Title.Text("Line roadmap"),
                  new ExcelChartDefinition.Title.Formula("B1"))));

      WorkbookReadResult.ChartsResult lineRead =
          assertInstanceOf(
              WorkbookReadResult.ChartsResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetCharts("charts", "Charts")));
      ExcelChartSnapshot initialLine = ExcelChartTestSupport.chart(lineRead.charts(), "OpsLine");
      assertEquals(new ExcelChartSnapshot.Title.Text("Line roadmap"), initialLine.title());

      sheet.setChart(
          lineChartDefinition(
              "OpsLine",
              ExcelChartTestSupport.anchor(6, 2, 12, 18),
              new ExcelChartDefinition.Title.Text("Line focus"),
              new ExcelChartDefinition.Title.Text("Actual")));
      ExcelChartSnapshot updatedLine = ExcelChartTestSupport.chart(sheet.charts(), "OpsLine");
      assertEquals(ExcelChartTestSupport.anchor(6, 2, 12, 18), updatedLine.anchor());
      assertEquals(new ExcelChartSnapshot.Title.Text("Line focus"), updatedLine.title());

      sheet.setChart(
          pieChartDefinition(
              "OpsPie",
              ExcelChartTestSupport.anchor(13, 1, 19, 12),
              new ExcelChartDefinition.Title.Text("Share"),
              90));
      ExcelChartSnapshot initialPie = ExcelChartTestSupport.chart(sheet.charts(), "OpsPie");
      assertEquals(
          90,
          ExcelChartTestSupport.singlePlot(initialPie, ExcelChartSnapshot.Pie.class)
              .firstSliceAngle());

      sheet.setChart(
          pieChartDefinition(
              "OpsPie",
              ExcelChartTestSupport.anchor(14, 2, 20, 13),
              new ExcelChartDefinition.Title.Text("Updated share"),
              120));
      ExcelChartSnapshot updatedPie = ExcelChartTestSupport.chart(sheet.charts(), "OpsPie");
      assertEquals(ExcelChartTestSupport.anchor(14, 2, 20, 13), updatedPie.anchor());
      assertEquals(
          120,
          ExcelChartTestSupport.singlePlot(updatedPie, ExcelChartSnapshot.Pie.class)
              .firstSliceAngle());

      IllegalArgumentException invalidChartTitle =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      lineChartDefinition(
                          "BadTitle",
                          ExcelChartTestSupport.anchor(28, 1, 34, 8),
                          new ExcelChartDefinition.Title.Formula("A2:A4"),
                          new ExcelChartDefinition.Title.Text("Plan"))));
      assertTrue(
          invalidChartTitle
              .getMessage()
              .contains("Chart title formula must resolve to a single cell"));

      IllegalArgumentException invalidSeriesTitle =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      lineChartDefinition(
                          "OpsLine",
                          ExcelChartTestSupport.anchor(8, 3, 14, 19),
                          new ExcelChartDefinition.Title.Text("Broken"),
                          new ExcelChartDefinition.Title.Formula("A2:A4"))));
      assertTrue(
          invalidSeriesTitle
              .getMessage()
              .contains("Series title formula must resolve to a single cell"));
    }
  }

  @Test
  void typeChangesSeriesRemovalAndFormulaSeriesTitlesStayDeterministic() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartTestSupport.seedChartNamedRanges(workbook, "Charts");

      sheet.setChart(
          barChartDefinition(
              "OpsBarSeries",
              ExcelChartTestSupport.anchor(1, 32, 7, 46),
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Text("Plan"),
                      ExcelChartTestSupport.ref("ChartCategories"),
                      ExcelChartTestSupport.ref("ChartPlan"),
                      null,
                      null,
                      null,
                      null),
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Formula("C1"),
                      ExcelChartTestSupport.ref("ChartCategories"),
                      ExcelChartTestSupport.ref("ChartActual"),
                      null,
                      null,
                      null,
                      null))));
      sheet.setChart(
          barChartDefinition(
              "OpsBarSeries",
              ExcelChartTestSupport.anchor(1, 32, 7, 46),
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.None(),
                      ExcelChartTestSupport.ref("ChartCategories"),
                      ExcelChartTestSupport.ref("ChartPlan"),
                      null,
                      null,
                      null,
                      null))));
      ExcelChartSnapshot barChart = ExcelChartTestSupport.chart(sheet.charts(), "OpsBarSeries");
      assertEquals(
          1,
          ExcelChartTestSupport.singlePlot(barChart, ExcelChartSnapshot.Bar.class).series().size());

      sheet.setChart(
          lineChartDefinition(
              "OpsType",
              ExcelChartTestSupport.anchor(8, 32, 14, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Title.Formula("B2")));
      sheet.setChart(
          pieChartDefinition(
              "OpsType",
              ExcelChartTestSupport.anchor(8, 32, 14, 46),
              new ExcelChartDefinition.Title.Text("Type change"),
              45));

      ExcelChartSnapshot updatedChart = ExcelChartTestSupport.chart(sheet.charts(), "OpsType");
      assertEquals(
          45,
          ExcelChartTestSupport.singlePlot(updatedChart, ExcelChartSnapshot.Pie.class)
              .firstSliceAngle());

      sheet.setChart(
          lineChartDefinition(
              "OpsFormula",
              ExcelChartTestSupport.anchor(8, 32, 14, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Title.Formula("B2")));
      ExcelChartSnapshot lineChart = ExcelChartTestSupport.chart(sheet.charts(), "OpsFormula");
      ExcelChartSnapshot.Series firstSeries =
          ExcelChartTestSupport.singlePlot(lineChart, ExcelChartSnapshot.Line.class)
              .series()
              .getFirst();
      ExcelChartSnapshot.Title.Formula title =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, firstSeries.title());
      assertEquals("10.0", title.cachedText());
    }
  }

  private static ExcelChartDefinition lineChartDefinition(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      ExcelChartDefinition.Title seriesTitle) {
    return ExcelChartTestSupport.lineChart(
        name,
        anchor,
        title,
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        List.of(
            new ExcelChartDefinition.Series(
                seriesTitle,
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartPlan"),
                null,
                null,
                null,
                null)));
  }

  private static ExcelChartDefinition pieChartDefinition(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      int firstSliceAngle) {
    return ExcelChartTestSupport.pieChart(
        name,
        anchor,
        title,
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.ZERO,
        false,
        true,
        firstSliceAngle,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Text("Actual"),
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartActual"),
                null,
                null,
                null,
                45L)));
  }

  private static ExcelChartDefinition barChartDefinition(
      String name, ExcelDrawingAnchor.TwoCell anchor, List<ExcelChartDefinition.Series> series) {
    return ExcelChartTestSupport.barChart(
        name,
        anchor,
        new ExcelChartDefinition.Title.None(),
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        ExcelChartBarDirection.COLUMN,
        series);
  }
}
