package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.assertThat;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.assertions;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.inspect;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.inspections;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingCoordinateReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookFactAssertion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.engine.poi.NativeInvocationContext;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Request assertions must use stored cache facts even when referenced live cells disagree. */
class StoredChartFactsWorkflowTest {
  @Test
  void requestInspectionAndAssertionsDistinguishStoredSevenFromLiveNinetyNine(
      @TempDir Path directory) throws IOException {
    Path source = directory.resolve("stored.xlsx");
    writeNativeFixture(source);
    var selector = new ChartSelector.AllOnSheet(new SheetName("Live"));
    var plan =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.toString(), Optional.empty()),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            assertions(
                assertThat(
                    "stored-facts",
                    selector,
                    new WorkbookFactAssertion.ChartFacts(List.of(expected("7.0"))))),
            inspections(
                inspect("read", selector, new WorkbookAssetIntrospectionQuery.GetCharts())));
    var success = assertInstanceOf(WorkbookResult.Success.class, execute(plan, directory));
    var read =
        assertInstanceOf(
            WorkbookAssetInspectionResult.ChartsResult.class, success.inspections().getFirst());
    assertEquals(List.of(expected("7.0")), read.charts());
    var negative =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.toString(), Optional.empty()),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            assertions(
                assertThat(
                    "live-is-not-cache",
                    selector,
                    new WorkbookFactAssertion.ChartFacts(List.of(expected("99.25"))))),
            List.of());
    var refused = assertInstanceOf(WorkbookResult.Failure.class, execute(negative, directory));
    assertEquals(
        dev.erst.gridgrind.core.problem.GridGrindProblemCode.ASSERTION_FAILED,
        refused.problem().code());
  }

  @Test
  void requestChartAndDrawingInspectionRetainNumericAxisCrossing(@TempDir Path directory)
      throws IOException {
    Path source = directory.resolve("crossing.xlsx");
    try (var workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      var chart =
          workbook
              .createSheet("Native")
              .createDrawingPatriarch()
              .createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 4, 4));
      chart.getGraphicFrame().setName("Numeric crossing");
      var data =
          (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData)
              chart.createData(
                  ChartTypes.BAR,
                  chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM),
                  chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT));
      data.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);
      data.setBarGrouping(org.apache.poi.xddf.usermodel.chart.BarGrouping.CLUSTERED);
      var axis = chart.getCTChart().getPlotArea().getValAxArray(0);
      axis.unsetCrosses();
      axis.unsetDelete();
      axis.addNewCrossesAt();
      try (var cursor = axis.getCrossesAt().newCursor()) {
        cursor.setAttributeText(new javax.xml.namespace.QName("val"), "2.500e0");
      }
      assertTrue(chart.getCTChartSpace().validate());
      workbook.write(output);
    }
    var zero = new DrawingCoordinateReport.Emu(0);
    var anchor =
        new DrawingAnchorReport.TwoCell(
            new DrawingMarkerReport(0, 0, zero, zero),
            new DrawingMarkerReport(4, 4, zero, zero),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    var axes =
        List.of(
            new ChartReport.Axis(
                dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind.CATEGORY,
                dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition.BOTTOM,
                new ChartReport.AxisCrossing.Mode(
                    dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses.AUTO_ZERO),
                true),
            new ChartReport.Axis(
                dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind.VALUE,
                dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition.LEFT,
                new ChartReport.AxisCrossing.Value("2.500e0"),
                true));
    var expected =
        new ChartReport(
            "Numeric crossing",
            anchor,
            new ChartReport.Title.None(),
            new ChartReport.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            List.of(
                new ChartReport.Bar(
                    false,
                    dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection.COLUMN,
                    dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping.CLUSTERED,
                    Optional.empty(),
                    Optional.empty(),
                    axes,
                    List.of())));
    var selector = new ChartSelector.AllOnSheet(new SheetName("Native"));
    var plan =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.toString(), Optional.empty()),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            assertions(
                assertThat(
                    "numeric-crossing",
                    selector,
                    new WorkbookFactAssertion.ChartFacts(List.of(expected)))),
            inspections(
                inspect("chart", selector, new WorkbookAssetIntrospectionQuery.GetCharts()),
                inspect(
                    "drawing",
                    new dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet(
                        new SheetName("Native")),
                    new WorkbookAssetIntrospectionQuery.GetDrawingObjects())));
    var success = assertInstanceOf(WorkbookResult.Success.class, execute(plan, directory));
    assertEquals(
        List.of(expected),
        assertInstanceOf(
                WorkbookAssetInspectionResult.ChartsResult.class, success.inspections().getFirst())
            .charts());
    assertEquals(
        1,
        assertInstanceOf(
                WorkbookAssetInspectionResult.DrawingObjectsResult.class,
                success.inspections().get(1))
            .drawingObjects()
            .size());
  }

  private static WorkbookResult execute(WorkbookPlan plan, Path directory) {
    return ExecutionContextFixtureSupport.execute(
        new DefaultGridGrindRequestExecutor(
            NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD)),
        plan,
        directory);
  }

  private static void writeNativeFixture(Path path) throws IOException {
    try (var workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(path)) {
      var sheet = workbook.createSheet("Live");
      sheet.createRow(0).createCell(0).setCellValue(99.25);
      var chart =
          sheet.createDrawingPatriarch().createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));
      chart.getGraphicFrame().setName("Stored");
      chart.createData(ChartTypes.PIE, null, null);
      chart.getCTChart().getPlotVisOnly().setVal(true);
      var series = chart.getCTChart().getPlotArea().getPieChartArray(0).addNewSer();
      series.addNewIdx().setVal(0);
      series.addNewOrder().setVal(0);
      var categories = series.addNewCat().addNewStrLit();
      categories.addNewPtCount().setVal(1);
      var category = categories.addNewPt();
      category.setIdx(0);
      category.setV("stored");
      var reference = series.addNewVal().addNewNumRef();
      reference.setF("Live!$A$1");
      var cache = reference.addNewNumCache();
      cache.addNewPtCount().setVal(1);
      var point = cache.addNewPt();
      point.setIdx(0);
      point.setV("7.0");
      assertTrue(chart.getCTChartSpace().validate());
      workbook.write(output);
    }
  }

  private static ChartReport expected(String storedValue) {
    var zero = new DrawingCoordinateReport.Emu(0);
    var anchor =
        new DrawingAnchorReport.TwoCell(
            new DrawingMarkerReport(0, 0, zero, zero),
            new DrawingMarkerReport(5, 5, zero, zero),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    var series =
        new ChartReport.Series(
            new ChartReport.Title.None(),
            new ChartReport.DataSource.StringLiteral(
                new ChartReport.StringCache.Stored(
                    Optional.of(1L),
                    List.of(new ChartReport.StringPoint(Optional.of(0L), Optional.of("stored"))))),
            new ChartReport.DataSource.NumericReference(
                Optional.of("Live!$A$1"),
                new ChartReport.NumericCache.Stored(
                    Optional.of(1L),
                    Optional.empty(),
                    List.of(
                        new ChartReport.NumericPoint(
                            Optional.of(0L), Optional.of(storedValue), Optional.empty())))),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    return new ChartReport(
        "Stored",
        anchor,
        new ChartReport.Title.None(),
        new ChartReport.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        List.of(new ChartReport.Pie(false, Optional.empty(), List.of(series))));
  }
}
