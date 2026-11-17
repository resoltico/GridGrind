package dev.erst.gridgrind.engine.charts;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Proxy;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;

/** Independent stored native fixtures challenge live resolution and lossy XDDF point access. */
class StoredChartReaderTest {
  @ParameterizedTest
  @ValueSource(
      strings = {
        "chart-stored-sparse.xlsx",
        "chart-stored-missing-cache.xlsx",
        "chart-stored-missing-declarations.xlsx"
      })
  void realChartAndDrawingReadsPreserveStoredFactsAndDoNotMutateXml(String fixture)
      throws IOException {
    try (var input =
            Objects.requireNonNull(getClass().getResourceAsStream("/workbooks/" + fixture));
        var workbook = new XSSFWorkbook(input)) {
      var sheet = workbook.getSheetAt(0);
      var chart = sheet.getDrawingPatriarch().getCharts().getFirst();
      String nativeBefore = chart.getCTChartSpace().xmlText();
      var adapter = new DrawingAdapter();
      var report = adapter.charts(sheet).getFirst();
      var plot = assertInstanceOf(ChartReport.Bar.class, report.plots().getFirst());
      var source =
          assertInstanceOf(
              ChartReport.DataSource.NumericReference.class, plot.series().getFirst().values());
      if ("chart-stored-missing-cache.xlsx".equals(fixture)) {
        assertEquals(Optional.of("'Missing source'!$B$2:$B$5"), source.formula());
        assertInstanceOf(ChartReport.NumericCache.Missing.class, source.cache());
      } else {
        var cache = assertInstanceOf(ChartReport.NumericCache.Stored.class, source.cache());
        boolean absentDeclarations = "chart-stored-missing-declarations.xlsx".equals(fixture);
        assertEquals(
            absentDeclarations ? Optional.empty() : Optional.of(9L), cache.declaredCount());
        assertEquals(Optional.of(""), cache.formatCode());
        assertEquals(
            List.of(
                new ChartReport.NumericPoint(
                    absentDeclarations ? Optional.empty() : Optional.of(4L),
                    Optional.of("#N/A"),
                    Optional.empty()),
                new ChartReport.NumericPoint(
                    Optional.of(0L), Optional.of("12.75"), Optional.empty()),
                new ChartReport.NumericPoint(Optional.of(2L), Optional.of(""), Optional.empty()),
                new ChartReport.NumericPoint(Optional.of(2L), Optional.of("NaN"), Optional.empty()),
                new ChartReport.NumericPoint(Optional.of(7L), Optional.empty(), Optional.empty())),
            cache.points());
      }
      var title = assertInstanceOf(ChartReport.Title.Formula.class, report.title());
      assertEquals(
          new ChartReport.StringCache.Stored(
              Optional.of(2L),
              List.of(
                  new ChartReport.StringPoint(Optional.of(1L), Optional.of("second")),
                  new ChartReport.StringPoint(Optional.of(0L), Optional.of("stored title")))),
          title.cache());
      var seriesTitle =
          assertInstanceOf(ChartReport.Title.Formula.class, plot.series().getFirst().title());
      assertEquals(
          new ChartReport.StringCache.Stored(
              Optional.of(2L),
              List.of(
                  new ChartReport.StringPoint(Optional.of(1L), Optional.of("second")),
                  new ChartReport.StringPoint(Optional.of(0L), Optional.of("stored series")))),
          seriesTitle.cache());
      var inventory =
          assertInstanceOf(
              DrawingObjectReport.Chart.class, adapter.drawingObjects(sheet).getFirst());
      assertEquals(report.title(), inventory.title());
      sheet.getRow(1).getCell(1).setCellValue(199.5);
      assertEquals(report, adapter.charts(sheet).getFirst());
      assertEquals(nativeBefore, chart.getCTChartSpace().xmlText());
    }
  }

  @Test
  void schemaValidZeroSeriesAndUnresolvedAxesRemainStoredPlotFacts() throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var drawing = workbook.createSheet("Empty native plot").createDrawingPatriarch();
      var chart = drawing.createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));
      chart.createData(
          ChartTypes.BAR,
          chart.createCategoryAxis(AxisPosition.BOTTOM),
          chart.createValueAxis(AxisPosition.LEFT));
      assertTrue(chart.getCTChartSpace().validate());
      var snapshot = ChartAdapter.snapshotChart(chart, chart.getGraphicFrame());
      var bar = assertInstanceOf(ChartReport.Bar.class, snapshot.plots().getFirst());
      assertEquals(List.of(), bar.series());
      assertEquals(2, bar.axes().size());
      var area = chart.getCTChart().getPlotArea();
      area.removeCatAx(0);
      area.removeValAx(0);
      assertTrue(chart.getCTChartSpace().validate());
      var serialized = new ByteArrayOutputStream();
      workbook.write(serialized);
      try (var reopened = new XSSFWorkbook(new ByteArrayInputStream(serialized.toByteArray()))) {
        var storedDrawing = reopened.getSheetAt(0).getDrawingPatriarch();
        storedDrawing.getShapes();
        var storedChart = storedDrawing.getCharts().getFirst();
        var unresolved =
            assertInstanceOf(
                ChartReport.Bar.class,
                ChartAdapter.snapshotChart(storedChart, storedChart.getGraphicFrame())
                    .plots()
                    .getFirst());
        assertEquals(List.of(), unresolved.axes());
        assertEquals(List.of(), unresolved.series());
      }
    }
  }

  @Test
  void blankStoredNameAndDefaultLegendReadWithoutChangingNativeXml() throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var drawing = workbook.createSheet("Native legend").createDrawingPatriarch();
      var chart = drawing.createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));
      chart.createData(
          ChartTypes.BAR,
          chart.createCategoryAxis(AxisPosition.BOTTOM),
          chart.createValueAxis(AxisPosition.LEFT));
      chart.getGraphicFrame().setName("");
      var legend = chart.getCTChart().addNewLegend();
      assertTrue(chart.getCTChartSpace().validate());
      String before = chart.getCTChartSpace().xmlText();
      var report = ChartAdapter.snapshotChart(chart, chart.getGraphicFrame());
      assertEquals("", report.name());
      assertEquals(
          new ChartReport.Legend.Visible(
              dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition.RIGHT),
          report.legend());
      assertEquals(
          "", ChartAdapter.snapshotChartDrawingObject(chart, chart.getGraphicFrame()).name());
      assertFalse(legend.isSetOverlay());
      assertEquals(before, chart.getCTChartSpace().xmlText());
    }
  }

  @ParameterizedTest
  @ValueSource(strings = {"1%", "5%", "9%", "90%"})
  void observesTheNativeHoleSizeRangeWithoutApplyingAuthoredBounds(String value)
      throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart =
          workbook
              .createSheet("Native doughnut")
              .createDrawingPatriarch()
              .createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));
      chart.createData(ChartTypes.DOUGHNUT, null, null);
      chart.getCTChart().getPlotArea().getDoughnutChartArray(0).addNewHoleSize().setVal(value);
      assertTrue(chart.getCTChartSpace().validate());
      var report =
          assertInstanceOf(
              ChartReport.Doughnut.class,
              ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()).plots().getFirst());
      assertEquals(
          Optional.of(Integer.valueOf(value.substring(0, value.length() - 1))), report.holeSize());
      chart.getCTChart().getPlotArea().getDoughnutChartArray(0).getHoleSize().setVal("0%");
      assertEquals(
          WorkbookContentIssue.INVALID_NATIVE_VALUE,
          assertThrows(
                  WorkbookContentRefusalException.class,
                  () -> ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()))
              .issue());
    }
  }

  @Test
  void literalStringAndNumericSourcesUseTheSameStoredDeclarationsWithoutAuthoringRestrictions() {
    var reader = new StoredChartReader(Optional.empty());
    var source = CTAxDataSource.Factory.newInstance();
    var strings = source.addNewStrLit();
    strings.addNewPtCount().setVal(7);
    var string = strings.addNewPt();
    string.setIdx(5);
    string.setV("");
    assertEquals(
        new ChartReport.DataSource.StringLiteral(
            new ChartReport.StringCache.Stored(
                Optional.of(7L),
                List.of(new ChartReport.StringPoint(Optional.of(5L), Optional.of(""))))),
        reader.source(source));
    source.unsetStrLit();
    var numbers = source.addNewNumLit();
    numbers.setFormatCode("");
    var number = numbers.addNewPt();
    number.setIdx(4);
    number.setV("#DIV/0!");
    number.setFormatCode("General");
    assertEquals(
        new ChartReport.DataSource.NumericLiteral(
            new ChartReport.NumericCache.Stored(
                Optional.empty(),
                Optional.of(""),
                List.of(
                    new ChartReport.NumericPoint(
                        Optional.of(4L), Optional.of("#DIV/0!"), Optional.of("General"))))),
        reader.source(source));
  }

  @Test
  void missingUnmodeledAndMalformedChoiceAreDistinctOutcomes() {
    var reader = new StoredChartReader(Optional.of("/xl/charts/fixture.xml"));
    var source = CTAxDataSource.Factory.newInstance();
    assertInstanceOf(ChartReport.DataSource.Missing.class, reader.source(source));
    source.addNewMultiLvlStrRef();
    assertEquals(
        "MULTI_LEVEL_STRING_REFERENCE",
        assertInstanceOf(ChartReport.DataSource.Unsupported.class, reader.source(source))
            .nativeType());
    source.addNewNumRef();
    var refusal = assertThrows(WorkbookContentRefusalException.class, () -> reader.source(source));
    assertEquals(WorkbookContentIssue.INVALID_NATIVE_VALUE, refusal.issue());
    assertEquals(
        Optional.of("/xl/charts/fixture.xml"), refusal.location().orElseThrow().partName());
  }

  @Test
  void invalidNativeUnsignedValueRefusesInsteadOfInventingAnIndex() {
    var source = CTAxDataSource.Factory.newInstance();
    var point = source.addNewNumLit().addNewPt();
    point.setIdx(-1);
    var refusal =
        assertThrows(
            WorkbookContentRefusalException.class,
            () -> new StoredChartReader(Optional.empty()).source(source));
    assertEquals(WorkbookContentIssue.INVALID_NATIVE_VALUE, refusal.issue());
    assertEquals(
        Optional.of("chart.numericCache.point.index"),
        refusal.location().orElseThrow().fieldName());
    assertFalse(refusal.getMessage().contains("-1"));
  }

  static Stream<RuntimeException> unexpectedNativeFailures() {
    return Stream.of(
        new IllegalArgumentException("private native argument"),
        new IllegalStateException("private native invariant"));
  }

  @ParameterizedTest
  @MethodSource("unexpectedNativeFailures")
  void unexpectedNativeExceptionsCannotBecomeUnsupportedFactsOrCallerInput(RuntimeException cause) {
    var source = failingNativeSource(cause);
    var wrapped =
        assertThrows(
            UnexpectedNativeContentException.class,
            () -> new StoredChartReader(Optional.empty()).source(source));
    assertSame(cause, wrapped.getCause());
  }

  @Test
  void fatalNativeFailureRetainsItsIdentity() {
    var fatal = new AssertionError("fatal native reader");
    var source = failingNativeSource(fatal);
    assertSame(
        fatal,
        assertThrows(
            AssertionError.class, () -> new StoredChartReader(Optional.empty()).source(source)));
  }

  // A native fault proxy must use its schema interface's defining loader, never an embedding TCCL.
  @SuppressWarnings("PMD.UseProperClassLoader")
  private static CTAxDataSource failingNativeSource(Throwable failure) {
    return (CTAxDataSource)
        Proxy.newProxyInstance(
            CTAxDataSource.class.getClassLoader(),
            new Class<?>[] {CTAxDataSource.class},
            (proxy, method, arguments) -> {
              throw failure;
            });
  }
}
