package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.DisplayBlanks;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** POI-authored interoperability and truthful unsupported-chart readback coverage. */
class ExcelChartInteropTest {
  @Test
  void unsupportedSinglePlotChartsReadBackAsUnsupportedPlots() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-unsupported-single-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      var anchor = drawing.createAnchor(0, 0, 0, 0, 4, 1, 11, 16);
      XSSFChart chart = drawing.createChart(anchor);
      chart.getGraphicFrame().setName("AreaOnly");
      XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
      valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
      var categories =
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4"));
      var values =
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4"));
      var areaData = chart.createData(ChartTypes.AREA, categoryAxis, valueAxis);
      areaData.addSeries(categories, values).setTitle("Plan", null);
      chart.plot(areaData);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelChartSnapshot areaChart = workbook.sheet("Charts").charts().getFirst();
      assertEquals("AreaOnly", areaChart.name());
      assertInstanceOf(
          ExcelChartSnapshot.Area.class,
          ExcelChartTestSupport.singlePlot(areaChart, ExcelChartSnapshot.Area.class));
    }
  }

  @Test
  void directPoiChartReadbackCoversLiteralSourcesTitlesAndAxisVariants() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-literal-readback-");
    writeLiteralChartReadbackWorkbook(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Charts");

      ExcelChartSnapshot cachedFormulaLine =
          sheet.charts().stream()
              .filter(snapshot -> snapshot.name().startsWith("Chart-"))
              .findFirst()
              .orElseThrow();
      ExcelChartSnapshot.Title.Formula cachedLineTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, cachedFormulaLine.title());
      assertEquals("Charts!$B$1", cachedLineTitle.formula());
      assertEquals("Plan", cachedLineTitle.cachedText());
      ExcelChartSnapshot.Line cachedLinePlot =
          ExcelChartTestSupport.singlePlot(cachedFormulaLine, ExcelChartSnapshot.Line.class);
      assertEquals(
          List.of(ExcelChartAxisPosition.TOP, ExcelChartAxisPosition.RIGHT),
          cachedLinePlot.axes().stream().map(ExcelChartSnapshot.Axis::position).toList());

      ExcelChartSnapshot.Series cachedLineSeries = cachedLinePlot.series().getFirst();
      ExcelChartSnapshot.Title.Formula cachedLineSeriesTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, cachedLineSeries.title());
      assertEquals("Actual", cachedLineSeriesTitle.cachedText());
      assertEquals(
          List.of("Jan", "Mar", ""),
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringLiteral.class, cachedLineSeries.categories())
              .values());

      ExcelChartSnapshot literalPie = ExcelChartTestSupport.chart(sheet.charts(), "LiteralPie");
      assertTrue(
          ExcelChartTestSupport.singlePlot(literalPie, ExcelChartSnapshot.Pie.class).varyColors());
      assertEquals(
          75,
          ExcelChartTestSupport.singlePlot(literalPie, ExcelChartSnapshot.Pie.class)
              .firstSliceAngle());

      ExcelDrawingObjectSnapshot.Chart uncachedFormulaDrawingObject =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> "FormulaNoCache".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals("Actual", uncachedFormulaDrawingObject.title());
    }
  }

  private static void writeLiteralChartReadbackWorkbook(Path workbookPath) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart cachedFormulaLine =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 8, 12));
      cachedFormulaLine.getGraphicFrame().setName("");
      var cachedLineTitle = cachedFormulaLine.getCTChart().addNewTitle();
      var cachedLineTitleRef = cachedLineTitle.addNewTx().addNewStrRef();
      cachedLineTitleRef.setF("Charts!$B$1");
      var cachedLineTitleCache = cachedLineTitleRef.addNewStrCache();
      cachedLineTitleCache.addNewPtCount().setVal(1);
      var cachedLineTitlePoint = cachedLineTitleCache.addNewPt();
      cachedLineTitlePoint.setIdx(0);
      cachedLineTitlePoint.setV("Plan");
      cachedFormulaLine
          .getOrAddLegend()
          .setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.LEFT);
      cachedFormulaLine.displayBlanksAs(DisplayBlanks.SPAN);
      XDDFCategoryAxis cachedLineCategoryAxis =
          cachedFormulaLine.createCategoryAxis(AxisPosition.TOP);
      XDDFValueAxis cachedLineValueAxis = cachedFormulaLine.createValueAxis(AxisPosition.RIGHT);
      cachedLineValueAxis.setCrosses(AxisCrosses.MAX);
      var cachedLineData =
          cachedFormulaLine.createData(
              ChartTypes.LINE, cachedLineCategoryAxis, cachedLineValueAxis);
      CTSerTx cachedLineSeriesTitle =
          ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series)
                  cachedLineData.addSeries(
                      XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Mar", null}),
                      XDDFDataSourcesFactory.fromArray(new Double[] {10d, 18d, 15d})))
              .getCTLineSer()
              .addNewTx();
      cachedLineSeriesTitle.addNewStrRef().setF("Charts!$C$1");
      cachedFormulaLine.plot(cachedLineData);

      XSSFChart uncachedFormulaLine =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 9, 1, 16, 12));
      uncachedFormulaLine.getGraphicFrame().setName("FormulaNoCache");
      var uncachedLineTitle = uncachedFormulaLine.getCTChart().addNewTitle();
      uncachedLineTitle.addNewTx().addNewStrRef().setF("Charts!$C$1");
      uncachedFormulaLine.displayBlanksAs(DisplayBlanks.ZERO);
      XDDFCategoryAxis uncachedLineCategoryAxis =
          uncachedFormulaLine.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis uncachedLineValueAxis = uncachedFormulaLine.createValueAxis(AxisPosition.LEFT);
      uncachedLineValueAxis.setCrosses(AxisCrosses.MIN);
      var uncachedLineData =
          uncachedFormulaLine.createData(
              ChartTypes.LINE, uncachedLineCategoryAxis, uncachedLineValueAxis);
      uncachedLineData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4")));
      uncachedFormulaLine.plot(uncachedLineData);

      XSSFChart literalPie = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 17, 1, 24, 12));
      literalPie.getGraphicFrame().setName("LiteralPie");
      literalPie.getCTChart().addNewTitle();
      literalPie
          .getOrAddLegend()
          .setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.TOP);
      var literalPieData =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData)
              literalPie.createData(ChartTypes.PIE, null, null);
      literalPieData.setVaryColors(true);
      literalPieData.setFirstSliceAngle(75);
      literalPieData.addSeries(
          XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb", "Mar"}),
          XDDFDataSourcesFactory.fromArray(new Double[] {12d, 16d, 21d}));
      literalPie.plot(literalPieData);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }
  }
}
