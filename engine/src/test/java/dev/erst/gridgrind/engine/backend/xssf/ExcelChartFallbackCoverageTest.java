package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.charts.ChartAdapter.AxisRegistry;
import dev.erst.gridgrind.engine.formulas.FormulaContext;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Focused coverage for chart helper fallback and no-runtime wrapper branches. */
class ExcelChartFallbackCoverageTest {
  @Test
  void sourceScalarWrappersAndRuntimeFallbacksStayDeterministic() throws IOException {
    assertEquals(
        new ChartAdapter.CellScalar(ChartAdapter.CellScalarKind.STRING, "hello", 0d),
        ChartAdapter.scalarFromFormula(new FormulaProbeCell(CellType.STRING, "hello", 0d, false)));
    assertEquals(
        new ChartAdapter.CellScalar(ChartAdapter.CellScalarKind.NUMERIC, null, 42d),
        ChartAdapter.scalarFromFormula(new FormulaProbeCell(CellType.NUMERIC, "", 42d, false)));
    assertEquals(
        new ChartAdapter.CellScalar(ChartAdapter.CellScalarKind.STRING, "true", 0d),
        ChartAdapter.scalarFromFormula(new FormulaProbeCell(CellType.BOOLEAN, "", 0d, true)));

    IllegalArgumentException cachedErrorFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ChartAdapter.scalarFromFormula(
                    new FormulaProbeCell(CellType.ERROR, "", 0d, false)));
    assertTrue(cachedErrorFailure.getMessage().contains("must not cache error values"));

    assertEquals(
        new ChartAdapter.CellScalar(ChartAdapter.CellScalarKind.STRING, "hello", 0d),
        ChartAdapter.scalarFromFormula(
            new FormulaProbeCell(CellType.STRING, "hello", 0d, false), null));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      sheet.getRow(1).createCell(3).setCellFormula("1+1");

      assertEquals(
          "18.0", ChartAdapter.scalarText(sheet, new CellReference("Charts", 2, 1, true, true)));
      assertEquals("18.0", ChartAdapter.scalarText(sheet, new CellReference(2, 1)));
      assertEquals(
          new ChartAdapter.CellScalar(ChartAdapter.CellScalarKind.STRING, "", 0d),
          ChartAdapter.scalarFromFormula(
              sheet.getRow(1).getCell(3),
              FormulaRuntimeTestDouble.nullEvaluation(
                  workbook.getCreationHelper().createFormulaEvaluator())));

      IllegalArgumentException unresolvedFormulaFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ChartAdapter.scalarFromFormula(
                      sheet.getRow(1).getCell(3), new EvaluatedTypeRuntime(CellType.FORMULA)));
      assertTrue(
          unresolvedFormulaFailure.getMessage().contains("must expose a cached scalar result"));
    }
  }

  @Test
  void titleFallbacksPreferLiveResolutionButReturnEmbeddedCachesWhenNeeded() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 8, 10));
      chart.getGraphicFrame().setName("FallbackChart");
      XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
      XDDFLineChartData lineData =
          (XDDFLineChartData) chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
      lineData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(
              sheet, org.apache.poi.ss.util.CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(
              sheet, org.apache.poi.ss.util.CellRangeAddress.valueOf("B2:B4")));
      chart.plot(lineData);

      assertEquals(
          "FallbackChart",
          ChartAdapter.snapshotChartDrawingObject(chart, chart.getGraphicFrame()).name());
      assertEquals("Plan", ChartAdapter.resolvedTitleFormulaText(chart, "B1"));

      XSSFChart untitledChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 10, 1, 16, 10));
      assertEquals("", ChartAdapter.cachedTitleText(untitledChart, "'Missing'!$A$1"));
      assertEquals("", ChartAdapter.resolvedTitleFormulaText(chart, "'Missing'!$A$1"));

      XSSFChart titleWithoutTxChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 17, 1, 23, 10));
      titleWithoutTxChart.getCTChart().addNewTitle();
      assertEquals("", ChartAdapter.cachedTitleText(titleWithoutTxChart, "'Missing'!$A$1"));

      XSSFChart titleWithoutReferenceChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 24, 1, 30, 10));
      titleWithoutReferenceChart.getCTChart().addNewTitle().addNewTx();
      assertEquals("", ChartAdapter.cachedTitleText(titleWithoutReferenceChart, "'Missing'!$A$1"));

      XSSFChart titleWithoutCacheChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 31, 1, 37, 10));
      titleWithoutCacheChart
          .getCTChart()
          .addNewTitle()
          .addNewTx()
          .addNewStrRef()
          .setF("'Missing'!$A$1");
      assertEquals("", ChartAdapter.cachedTitleText(titleWithoutCacheChart, "'Missing'!$A$1"));

      XSSFChart titleWithEmptyCacheChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 38, 1, 44, 10));
      var emptyCacheRef =
          titleWithEmptyCacheChart.getCTChart().addNewTitle().addNewTx().addNewStrRef();
      emptyCacheRef.setF("'Missing'!$A$1");
      emptyCacheRef.addNewStrCache();
      assertEquals("", ChartAdapter.cachedTitleText(titleWithEmptyCacheChart, "'Missing'!$A$1"));

      var titleRef = chart.getCTChart().addNewTitle().addNewTx().addNewStrRef();
      titleRef.setF("'Missing'!$A$1");
      titleRef.addNewStrCache().addNewPtCount().setVal(1);
      var titlePoint = titleRef.getStrCache().addNewPt();
      titlePoint.setIdx(0);
      titlePoint.setV("cached-title");

      assertEquals("cached-title", ChartAdapter.cachedTitleText(chart, "'Missing'!$A$1"));

      CTSerTx seriesTitle = CTSerTx.Factory.newInstance();
      var seriesRef = seriesTitle.addNewStrRef();
      seriesRef.setF("'Missing'!$B$1");
      seriesRef.addNewStrCache().addNewPtCount().setVal(1);
      var seriesPoint = seriesRef.getStrCache().addNewPt();
      seriesPoint.setIdx(0);
      seriesPoint.setV("cached-series");

      ChartReport.Title.Formula fallbackSeriesTitle =
          (ChartReport.Title.Formula) ChartAdapter.snapshotSeriesTitle(sheet, seriesTitle, null);
      assertEquals("'Missing'!$B$1", fallbackSeriesTitle.formula());
      assertEquals("cached-series", fallbackSeriesTitle.cachedText());
    }
  }

  @Test
  void snapshottingUsesExplicitGraphicFrameContextWithoutPoiReflection() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 8, 10));
      chart.getGraphicFrame().setName("DetachedFrame");
      chart.setTitleFormula("B1");
      XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
      XDDFLineChartData lineData =
          (XDDFLineChartData) chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
      lineData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(
              sheet, org.apache.poi.ss.util.CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(
              sheet, org.apache.poi.ss.util.CellRangeAddress.valueOf("B2:B4")));
      chart.plot(lineData);

      var graphicFrame = chart.getGraphicFrame();
      assertEquals(Optional.empty(), ChartAdapter.contextSheet((XSSFGraphicFrame) null));
      assertNull(ChartAdapter.contextSheet(null, null));
      assertEquals(sheet, ChartAdapter.contextSheet(chart, null));
      assertEquals(sheet, ChartAdapter.contextSheet(null, graphicFrame));
      assertEquals(
          Optional.empty(), ChartAdapter.optionalResolvedTitleFormulaText(null, null, "B1", null));
      assertEquals(
          Optional.of("Plan"),
          ChartAdapter.optionalResolvedTitleFormulaText(chart, null, "B1", null));
      assertEquals(
          Optional.of("Plan"),
          ChartAdapter.optionalResolvedTitleFormulaText(null, graphicFrame, "B1", null));
      assertEquals("Plan", ChartAdapter.cachedTitleText(chart, graphicFrame, "B1", null));

      ChartReport snapshot = ChartAdapter.snapshotChart(chart, graphicFrame);
      assertEquals("DetachedFrame", snapshot.name());
      assertInstanceOf(ChartReport.Line.class, snapshot.plots().getFirst());
      ChartReport.Title.Formula resolvedTitle =
          assertInstanceOf(ChartReport.Title.Formula.class, snapshot.title());
      assertEquals("B1", resolvedTitle.formula());
      assertEquals("Plan", resolvedTitle.cachedText());
    }
  }

  @Test
  void titleFormulaFallbackTreatsErrorCachedScalarsAsRecoverable() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      sheet.createRow(0).createCell(0).setCellValue("Labels");
      sheet.getRow(0).createCell(1).setCellFormula("1/0");
      workbook
          .getCreationHelper()
          .createFormulaEvaluator()
          .evaluateFormulaCell(sheet.getRow(0).getCell(1));

      XSSFChart chart =
          sheet
              .createDrawingPatriarch()
              .createChart(sheet.createDrawingPatriarch().createAnchor(0, 0, 0, 0, 1, 1, 8, 10));

      assertEquals(
          Optional.empty(), ChartAdapter.optionalResolvedTitleFormulaText(chart, null, "B1", null));
      assertEquals("", ChartAdapter.resolvedTitleFormulaText(chart, "B1"));
    }
  }

  @Test
  void titleFormulaFallbackReturnsEmptyForMalformedReferencesAndPropagatesRuntimeFailures()
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      sheet.createRow(0).createCell(0).setCellValue("Plan");
      sheet.getRow(0).createCell(1).setCellFormula("UPPER(A1)");

      XSSFChart chart =
          sheet
              .createDrawingPatriarch()
              .createChart(sheet.createDrawingPatriarch().createAnchor(0, 0, 0, 0, 1, 1, 8, 10));

      assertEquals(
          Optional.empty(),
          ChartAdapter.optionalResolvedTitleFormulaText(chart, null, "not-a-cell-reference", null));
      IllegalStateException runtimeFailure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ChartAdapter.optionalResolvedTitleFormulaText(
                      chart,
                      null,
                      "B1",
                      FormulaRuntimeTestDouble.alwaysFail(new IllegalStateException("boom"))));
      assertEquals("boom", runtimeFailure.getMessage());
    }
  }

  @Test
  void plotCreationWrapperDelegatesWithoutAnExplicitRuntime() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      XSSFChart chart =
          sheet
              .createDrawingPatriarch()
              .createChart(sheet.createDrawingPatriarch().createAnchor(0, 0, 0, 0, 1, 1, 8, 10));
      AxisRegistry axisRegistry = new AxisRegistry(chart);

      ChartAdapter.createPlot(
          sheet,
          chart,
          axisRegistry,
          new ResolvedChart.Line(
              false,
              ExcelChartGrouping.STANDARD,
              axes(),
              List.of(
                  new ResolvedChart.Series(
                      null,
                      ExcelChartTestSupport.ref("A2:A4"),
                      ExcelChartTestSupport.ref("B2:B4"),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()))));

      assertEquals(1, chart.getChartSeries().size());
      assertEquals(1, chart.getChartSeries().getFirst().getSeriesCount());
    }
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

  /**
   * Minimal runtime that forces one evaluated type for scalarFromFormula runtime branch coverage.
   */
  private static final class EvaluatedTypeRuntime implements FormulaRuntime {
    private final CellType evaluatedType;

    private EvaluatedTypeRuntime(CellType evaluatedType) {
      this.evaluatedType = evaluatedType;
    }

    @Override
    public CellValue evaluate(Cell cell) {
      throw new UnsupportedOperationException();
    }

    @Override
    public CellType evaluateFormulaCell(Cell cell) {
      return evaluatedType;
    }

    @Override
    public void clearCachedResults() {}

    @Override
    public String displayValue(DataFormatter formatter, Cell cell) {
      throw new UnsupportedOperationException();
    }

    @Override
    public FormulaContext context() {
      return FormulaContext.from(ResolvedFormulaEnvironment.defaults());
    }
  }
}
