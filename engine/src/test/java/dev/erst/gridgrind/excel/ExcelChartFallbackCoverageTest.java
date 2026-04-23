package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.VarHandle;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellBase;
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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Focused coverage for chart helper fallback and no-runtime wrapper branches. */
class ExcelChartFallbackCoverageTest {
  private static final VarHandle CHART_FRAME_HANDLE = chartFrameHandle();

  @Test
  void sourceScalarWrappersAndRuntimeFallbacksStayDeterministic() throws IOException {
    assertEquals(
        new ExcelDrawingController.CellScalar(
            ExcelDrawingController.CellScalarKind.STRING, "hello", 0d),
        ExcelChartSourceSupport.scalarFromFormula(
            new FormulaProbeCell(CellType.STRING, "hello", 0d, false)));
    assertEquals(
        new ExcelDrawingController.CellScalar(
            ExcelDrawingController.CellScalarKind.NUMERIC, null, 42d),
        ExcelChartSourceSupport.scalarFromFormula(
            new FormulaProbeCell(CellType.NUMERIC, "", 42d, false)));
    assertEquals(
        new ExcelDrawingController.CellScalar(
            ExcelDrawingController.CellScalarKind.STRING, "true", 0d),
        ExcelChartSourceSupport.scalarFromFormula(
            new FormulaProbeCell(CellType.BOOLEAN, "", 0d, true)));

    IllegalArgumentException cachedErrorFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelChartSourceSupport.scalarFromFormula(
                    new FormulaProbeCell(CellType.ERROR, "", 0d, false)));
    assertTrue(cachedErrorFailure.getMessage().contains("must not cache error values"));

    assertEquals(
        new ExcelDrawingController.CellScalar(
            ExcelDrawingController.CellScalarKind.STRING, "hello", 0d),
        ExcelChartSourceSupport.scalarFromFormula(
            new FormulaProbeCell(CellType.STRING, "hello", 0d, false), null));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      sheet.getRow(1).createCell(3).setCellFormula("1+1");

      assertEquals(
          "18.0",
          ExcelChartSourceSupport.scalarText(sheet, new CellReference("Charts", 2, 1, true, true)));
      assertEquals(
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING, "", 0d),
          ExcelChartSourceSupport.scalarFromFormula(
              sheet.getRow(1).getCell(3),
              FormulaRuntimeTestDouble.nullEvaluation(
                  workbook.getCreationHelper().createFormulaEvaluator())));

      IllegalArgumentException unresolvedFormulaFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelChartSourceSupport.scalarFromFormula(
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
          ExcelChartSnapshotSupport.snapshotChartDrawingObject(chart, chart.getGraphicFrame())
              .name());
      assertEquals("Plan", ExcelChartSnapshotSupport.resolvedTitleFormulaText(chart, "B1"));

      XSSFChart untitledChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 10, 1, 16, 10));
      assertEquals("", ExcelChartSnapshotSupport.cachedTitleText(untitledChart, "'Missing'!$A$1"));
      assertEquals("", ExcelChartSnapshotSupport.resolvedTitleFormulaText(chart, "'Missing'!$A$1"));

      XSSFChart titleWithoutTxChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 17, 1, 23, 10));
      titleWithoutTxChart.getCTChart().addNewTitle();
      assertEquals(
          "", ExcelChartSnapshotSupport.cachedTitleText(titleWithoutTxChart, "'Missing'!$A$1"));

      XSSFChart titleWithoutReferenceChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 24, 1, 30, 10));
      titleWithoutReferenceChart.getCTChart().addNewTitle().addNewTx();
      assertEquals(
          "",
          ExcelChartSnapshotSupport.cachedTitleText(titleWithoutReferenceChart, "'Missing'!$A$1"));

      XSSFChart titleWithoutCacheChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 31, 1, 37, 10));
      titleWithoutCacheChart
          .getCTChart()
          .addNewTitle()
          .addNewTx()
          .addNewStrRef()
          .setF("'Missing'!$A$1");
      assertEquals(
          "", ExcelChartSnapshotSupport.cachedTitleText(titleWithoutCacheChart, "'Missing'!$A$1"));

      XSSFChart titleWithEmptyCacheChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 38, 1, 44, 10));
      var emptyCacheRef =
          titleWithEmptyCacheChart.getCTChart().addNewTitle().addNewTx().addNewStrRef();
      emptyCacheRef.setF("'Missing'!$A$1");
      emptyCacheRef.addNewStrCache();
      assertEquals(
          "",
          ExcelChartSnapshotSupport.cachedTitleText(titleWithEmptyCacheChart, "'Missing'!$A$1"));

      var titleRef = chart.getCTChart().addNewTitle().addNewTx().addNewStrRef();
      titleRef.setF("'Missing'!$A$1");
      titleRef.addNewStrCache().addNewPtCount().setVal(1);
      var titlePoint = titleRef.getStrCache().addNewPt();
      titlePoint.setIdx(0);
      titlePoint.setV("cached-title");

      assertEquals(
          "cached-title", ExcelChartSnapshotSupport.cachedTitleText(chart, "'Missing'!$A$1"));

      CTSerTx seriesTitle = CTSerTx.Factory.newInstance();
      var seriesRef = seriesTitle.addNewStrRef();
      seriesRef.setF("'Missing'!$B$1");
      seriesRef.addNewStrCache().addNewPtCount().setVal(1);
      var seriesPoint = seriesRef.getStrCache().addNewPt();
      seriesPoint.setIdx(0);
      seriesPoint.setV("cached-series");

      ExcelChartSnapshot.Title.Formula fallbackSeriesTitle =
          (ExcelChartSnapshot.Title.Formula)
              ExcelChartSnapshotSupport.snapshotSeriesTitle(sheet, seriesTitle, null);
      assertEquals("'Missing'!$B$1", fallbackSeriesTitle.formula());
      assertEquals("cached-series", fallbackSeriesTitle.cachedText());
    }
  }

  @Test
  void snapshottingUsesExplicitGraphicFrameWhenPoiBackpointerIsMissing() throws IOException {
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
      detachGraphicFrame(chart);
      assertEquals("", ExcelChartSnapshotSupport.cachedTitleText(chart, "B1"));

      ExcelChartSnapshot snapshot = ExcelChartSnapshotSupport.snapshotChart(chart, graphicFrame);
      assertEquals("DetachedFrame", snapshot.name());
      assertInstanceOf(ExcelChartSnapshot.Line.class, snapshot.plots().getFirst());
      ExcelChartSnapshot.Title.Formula resolvedTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, snapshot.title());
      assertEquals("B1", resolvedTitle.formula());
      assertEquals("Plan", resolvedTitle.cachedText());
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
      ExcelChartAxisRegistry axisRegistry = new ExcelChartAxisRegistry(chart);

      ExcelChartPlotMutationSupport.createPlot(
          sheet,
          chart,
          axisRegistry,
          new ExcelChartDefinition.Line(
              false,
              ExcelChartGrouping.STANDARD,
              axes(),
              List.of(
                  new ExcelChartDefinition.Series(
                      null,
                      ExcelChartTestSupport.ref("A2:A4"),
                      ExcelChartTestSupport.ref("B2:B4"),
                      null,
                      null,
                      null,
                      null))));

      assertEquals(1, chart.getChartSeries().size());
      assertEquals(1, chart.getChartSeries().getFirst().getSeriesCount());
    }
  }

  private static List<ExcelChartDefinition.Axis> axes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static void detachGraphicFrame(XSSFChart chart) {
    CHART_FRAME_HANDLE.set(chart, null);
  }

  private static VarHandle chartFrameHandle() {
    try {
      return MethodHandles.privateLookupIn(XSSFChart.class, MethodHandles.lookup())
          .findVarHandle(
              XSSFChart.class, "frame", org.apache.poi.xssf.usermodel.XSSFGraphicFrame.class);
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("Failed to access POI chart graphic-frame backpointer", exception);
    }
  }

  /** Minimal formula cell stub for probing scalar-decoding branches without reflection. */
  private static final class FormulaProbeCell extends CellBase {
    private final CellType cachedFormulaResultType;
    private final String stringValue;
    private final double numericValue;
    private final boolean booleanValue;

    private FormulaProbeCell(
        CellType cachedFormulaResultType,
        String stringValue,
        double numericValue,
        boolean booleanValue) {
      this.cachedFormulaResultType = cachedFormulaResultType;
      this.stringValue = stringValue;
      this.numericValue = numericValue;
      this.booleanValue = booleanValue;
    }

    @Override
    public CellType getCellType() {
      return CellType.FORMULA;
    }

    @Override
    public CellType getCachedFormulaResultType() {
      return cachedFormulaResultType;
    }

    @Override
    public int getColumnIndex() {
      return 0;
    }

    @Override
    public int getRowIndex() {
      return 0;
    }

    @Override
    public org.apache.poi.ss.usermodel.Sheet getSheet() {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.Row getRow() {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getCellFormula() {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getStringCellValue() {
      return stringValue;
    }

    @Override
    public double getNumericCellValue() {
      return numericValue;
    }

    @Override
    public boolean getBooleanCellValue() {
      return booleanValue;
    }

    @Override
    public java.util.Date getDateCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public java.time.LocalDateTime getLocalDateTimeCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.RichTextString getRichStringCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public byte getErrorCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellValue(boolean value) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellErrorValue(byte value) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellStyle(org.apache.poi.ss.usermodel.CellStyle style) {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.CellStyle getCellStyle() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setAsActiveCell() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellComment(org.apache.poi.ss.usermodel.Comment comment) {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.Comment getCellComment() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void removeCellComment() {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.Hyperlink getHyperlink() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setHyperlink(org.apache.poi.ss.usermodel.Hyperlink link) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void removeHyperlink() {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellTypeImpl(CellType cellType) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellFormulaImpl(String formula) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void removeFormulaImpl() {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(double value) {
      throw new UnsupportedOperationException();
    }

    @Override
    @SuppressWarnings("PMD.ReplaceJavaUtilDate")
    protected void setCellValueImpl(java.util.Date value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(java.time.LocalDateTime value) {
      throw new UnsupportedOperationException();
    }

    @Override
    @SuppressWarnings("PMD.ReplaceJavaUtilCalendar")
    protected void setCellValueImpl(java.util.Calendar value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(String value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(org.apache.poi.ss.usermodel.RichTextString value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected SpreadsheetVersion getSpreadsheetVersion() {
      return SpreadsheetVersion.EXCEL2007;
    }

    @Override
    public org.apache.poi.ss.util.CellRangeAddress getArrayFormulaRange() {
      throw new UnsupportedOperationException();
    }

    @Override
    public boolean isPartOfArrayFormulaGroup() {
      return false;
    }

    @Override
    public void setBlank() {
      throw new UnsupportedOperationException();
    }
  }

  /**
   * Minimal runtime that forces one evaluated type for scalarFromFormula runtime branch coverage.
   */
  private static final class EvaluatedTypeRuntime implements ExcelFormulaRuntime {
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
    public ExcelFormulaRuntimeContext context() {
      return ExcelFormulaEnvironment.defaults().runtimeContext();
    }
  }
}
