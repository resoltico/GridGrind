package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** End-to-end mutation and introspection coverage for authored charts. */
class ExcelChartMutationFlowTest {
  @Test
  void lineAndPieChartsExerciseMutationExecutorAndIntrospection() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartTestSupport.seedChartNamedRanges(workbook, "Charts");

      executor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Charts")),
              new ResolvedMutationAction.SetChart(
                  lineChartDefinition(
                      "OpsLine",
                      ExcelChartTestSupport.anchor(4, 1, 10, 16),
                      new ResolvedChartTitle.Text("Line roadmap"),
                      new ChartTitleSpec.Formula("B1")))));

      dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult lineRead =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult.class,
              introspector.execute(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                      new StepId("charts"),
                      new ChartSelector.AllOnSheet(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("Charts")),
                      new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                          .GetCharts())));
      ChartReport initialLine = ExcelChartTestSupport.chart(lineRead.charts(), "OpsLine");
      assertEquals(new ChartReport.Title.Text("Line roadmap"), initialLine.title());

      sheet
          .drawings()
          .setChart(
              lineChartDefinition(
                  "OpsLine",
                  ExcelChartTestSupport.anchor(6, 2, 12, 18),
                  new ResolvedChartTitle.Text("Line focus"),
                  new ResolvedChartTitle.Text("Actual")));
      ChartReport updatedLine = ExcelChartTestSupport.chart(sheet.drawings().charts(), "OpsLine");
      assertEquals(
          DrawingGeometryFixtures.report(ExcelChartTestSupport.anchor(6, 2, 12, 18)),
          updatedLine.anchor());
      assertEquals(new ChartReport.Title.Text("Line focus"), updatedLine.title());

      sheet
          .drawings()
          .setChart(
              pieChartDefinition(
                  "OpsPie",
                  ExcelChartTestSupport.anchor(13, 1, 19, 12),
                  new ResolvedChartTitle.Text("Share"),
                  90));
      ChartReport initialPie = ExcelChartTestSupport.chart(sheet.drawings().charts(), "OpsPie");
      assertEquals(
          Optional.of(90),
          ExcelChartTestSupport.singlePlot(initialPie, ChartReport.Pie.class).firstSliceAngle());

      sheet
          .drawings()
          .setChart(
              pieChartDefinition(
                  "OpsPie",
                  ExcelChartTestSupport.anchor(14, 2, 20, 13),
                  new ResolvedChartTitle.Text("Updated share"),
                  120));
      ChartReport updatedPie = ExcelChartTestSupport.chart(sheet.drawings().charts(), "OpsPie");
      assertEquals(
          DrawingGeometryFixtures.report(ExcelChartTestSupport.anchor(14, 2, 20, 13)),
          updatedPie.anchor());
      assertEquals(
          Optional.of(120),
          ExcelChartTestSupport.singlePlot(updatedPie, ChartReport.Pie.class).firstSliceAngle());

      dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult exactChartRead =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult.class,
              introspector.execute(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                      new StepId("chart-by-name"),
                      new ChartSelector.ByName(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("Charts"), "OpsPie"),
                      new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                          .GetCharts())));
      assertEquals(1, exactChartRead.charts().size());
      assertEquals("OpsPie", exactChartRead.charts().getFirst().name());

      IllegalArgumentException invalidChartTitle =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet
                      .drawings()
                      .setChart(
                          lineChartDefinition(
                              "BadTitle",
                              ExcelChartTestSupport.anchor(28, 1, 34, 8),
                              new ChartTitleSpec.Formula("A2:A4"),
                              new ResolvedChartTitle.Text("Plan"))));
      assertTrue(
          invalidChartTitle
              .getMessage()
              .contains("Chart title formula must resolve to a single cell"));

      IllegalArgumentException invalidSeriesTitle =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet
                      .drawings()
                      .setChart(
                          lineChartDefinition(
                              "OpsLine",
                              ExcelChartTestSupport.anchor(8, 3, 14, 19),
                              new ResolvedChartTitle.Text("Broken"),
                              new ChartTitleSpec.Formula("A2:A4"))));
      assertTrue(
          invalidSeriesTitle
              .getMessage()
              .contains("Series title formula must resolve to a single cell"));
    }
  }

  @Test
  void typeChangesSeriesRemovalAndFormulaSeriesTitlesStayDeterministic() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartTestSupport.seedChartNamedRanges(workbook, "Charts");

      sheet
          .drawings()
          .setChart(
              barChartDefinition(
                  "OpsBarSeries",
                  ExcelChartTestSupport.anchor(1, 32, 7, 46),
                  List.of(
                      new ResolvedChart.Series(
                          new ResolvedChartTitle.Text("Plan"),
                          ExcelChartTestSupport.ref("ChartCategories"),
                          ExcelChartTestSupport.ref("ChartPlan"),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()),
                      new ResolvedChart.Series(
                          new ChartTitleSpec.Formula("C1"),
                          ExcelChartTestSupport.ref("ChartCategories"),
                          ExcelChartTestSupport.ref("ChartActual"),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))));
      sheet
          .drawings()
          .setChart(
              barChartDefinition(
                  "OpsBarSeries",
                  ExcelChartTestSupport.anchor(1, 32, 7, 46),
                  List.of(
                      new ResolvedChart.Series(
                          new ChartTitleSpec.None(),
                          ExcelChartTestSupport.ref("ChartCategories"),
                          ExcelChartTestSupport.ref("ChartPlan"),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))));
      ChartReport barChart = ExcelChartTestSupport.chart(sheet.drawings().charts(), "OpsBarSeries");
      assertEquals(
          1, ExcelChartTestSupport.singlePlot(barChart, ChartReport.Bar.class).series().size());

      sheet
          .drawings()
          .setChart(
              lineChartDefinition(
                  "OpsType",
                  ExcelChartTestSupport.anchor(8, 32, 14, 46),
                  new ChartTitleSpec.None(),
                  new ChartTitleSpec.Formula("B2")));
      sheet
          .drawings()
          .setChart(
              pieChartDefinition(
                  "OpsType",
                  ExcelChartTestSupport.anchor(8, 32, 14, 46),
                  new ResolvedChartTitle.Text("Type change"),
                  45));

      ChartReport updatedChart = ExcelChartTestSupport.chart(sheet.drawings().charts(), "OpsType");
      assertEquals(
          Optional.of(45),
          ExcelChartTestSupport.singlePlot(updatedChart, ChartReport.Pie.class).firstSliceAngle());

      sheet
          .drawings()
          .setChart(
              lineChartDefinition(
                  "OpsFormula",
                  ExcelChartTestSupport.anchor(8, 32, 14, 46),
                  new ChartTitleSpec.None(),
                  new ChartTitleSpec.Formula("B2")));
      ChartReport lineChart = ExcelChartTestSupport.chart(sheet.drawings().charts(), "OpsFormula");
      ChartReport.Series firstSeries =
          ExcelChartTestSupport.singlePlot(lineChart, ChartReport.Line.class).series().getFirst();
      ChartReport.Title.Formula title =
          assertInstanceOf(ChartReport.Title.Formula.class, firstSeries.title());
      assertEquals("10.0", title.cachedText());
    }
  }

  private static ResolvedChart lineChartDefinition(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      ResolvedChartTitle title,
      ResolvedChartTitle seriesTitle) {
    return ExcelChartTestSupport.lineChart(
        name,
        anchor,
        title,
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        List.of(
            new ResolvedChart.Series(
                seriesTitle,
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartPlan"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())));
  }

  private static ResolvedChart pieChartDefinition(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      ResolvedChartTitle title,
      int firstSliceAngle) {
    return ExcelChartTestSupport.pieChart(
        name,
        anchor,
        title,
        new ChartLegendSpec.Hidden(),
        ExcelChartDisplayBlanksAs.ZERO,
        false,
        true,
        firstSliceAngle,
        List.of(
            new ResolvedChart.Series(
                new ResolvedChartTitle.Text("Actual"),
                ExcelChartTestSupport.ref("ChartCategories"),
                ExcelChartTestSupport.ref("ChartActual"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(45L))));
  }

  private static ResolvedChart barChartDefinition(
      String name, DrawingAnchorSpec.TwoCell anchor, List<ResolvedChart.Series> series) {
    return ExcelChartTestSupport.barChart(
        name,
        anchor,
        new ChartTitleSpec.None(),
        new ChartLegendSpec.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        ExcelChartBarDirection.COLUMN,
        series);
  }
}
