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
  void snapshotUsesExplicitGraphicFrameAndPreservesMissingTitleCache() throws IOException {
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

      ChartReport snapshot = ChartAdapter.snapshotChart(chart, graphicFrame);
      assertEquals("DetachedFrame", snapshot.name());
      assertInstanceOf(ChartReport.Line.class, snapshot.plots().getFirst());
      ChartReport.Title.Formula resolvedTitle =
          assertInstanceOf(ChartReport.Title.Formula.class, snapshot.title());
      assertEquals(Optional.of("B1"), resolvedTitle.formula());
      assertInstanceOf(ChartReport.StringCache.Missing.class, resolvedTitle.cache());
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
