package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.charset.StandardCharsets;
import java.util.Base64;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Drawing chart-title and vary-color helper coverage. */
class ExcelDrawingChartHelperCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void drawingControllerReflectiveTitleNameAndVaryColorHelpersCoverResidualBranches()
      throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      workbook.createSheet("Other");
      sheet.createRow(0).createCell(0).setCellValue("Fallback title");

      assertThrows(
          IllegalArgumentException.class,
          () ->
              ExcelDrawingChartSupport.requiredDefinedNameFormula(
                  new DefinedNameStub("BlankStringSource", " ", -1)));

      org.apache.poi.ss.usermodel.Name workbookScoped = workbook.createName();
      workbookScoped.setNameName("ScopedSource");
      workbookScoped.setRefersToFormula("Ops!$A$1");
      assertSame(
          workbookScoped,
          ExcelDrawingChartSupport.resolveDefinedNameReference(sheet, "ScopedSource"));
      org.apache.poi.ss.usermodel.Name otherSheetScoped = workbook.createName();
      otherSheetScoped.setNameName("OtherOnly");
      otherSheetScoped.setSheetIndex(workbook.getSheetIndex("Other"));
      otherSheetScoped.setRefersToFormula("Other!$A$1");
      assertNull(ExcelDrawingChartSupport.resolveDefinedNameReference(sheet, "OtherOnly"));
      assertNull(ExcelDrawingChartSupport.resolveDefinedNameReference(sheet, "Bad-1"));

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      org.apache.poi.xssf.usermodel.XSSFChart firstChart =
          drawing.createChart(poiAnchor(drawing, 0, 0, 4, 6));
      firstChart.getGraphicFrame().setName("First");
      org.apache.poi.xssf.usermodel.XSSFChart secondChart =
          drawing.createChart(poiAnchor(drawing, 5, 0, 9, 6));
      secondChart.getGraphicFrame().setName("Second");
      assertSame(
          secondChart,
          ExcelDrawingChartSupport.chartForGraphicFrame(drawing, secondChart.getGraphicFrame()));

      org.apache.poi.xssf.usermodel.XSSFChart blankTitleChart =
          drawing.createChart(poiAnchor(drawing, 10, 0, 14, 6));
      blankTitleChart.setTitleText(" ");
      assertNotNull(blankTitleChart.getTitleText());
      assertTrue(blankTitleChart.getTitleText().getString().isBlank());
      assertInstanceOf(
          ExcelChartSnapshot.Title.None.class,
          ExcelDrawingChartSupport.snapshotTitle(blankTitleChart));
      org.apache.poi.xssf.usermodel.XSSFChart missingTextChart =
          drawing.createChart(poiAnchor(drawing, 15, 0, 19, 6));
      missingTextChart.getCTChart().addNewTitle().addNewTx();
      assertInstanceOf(
          ExcelChartSnapshot.Title.None.class,
          ExcelDrawingChartSupport.snapshotTitle(missingTextChart));
      org.apache.poi.xssf.usermodel.XSSFChart textTitleChart =
          drawing.createChart(poiAnchor(drawing, 15, 7, 19, 13));
      textTitleChart.setTitleText("Solid");
      assertEquals(
          new ExcelChartSnapshot.Title.Text("Solid"),
          ExcelDrawingChartSupport.snapshotTitle(textTitleChart));

      org.apache.poi.xssf.usermodel.XSSFChart cachedTitleChart =
          drawing.createChart(poiAnchor(drawing, 20, 0, 24, 6));
      assertEquals("", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$99"));
      cachedTitleChart.getCTChart().addNewTitle();
      assertEquals("", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$99"));
      cachedTitleChart.getCTChart().getTitle().addNewTx();
      assertEquals("", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$99"));
      cachedTitleChart.getCTChart().getTitle().getTx().addNewStrRef();
      cachedTitleChart.getCTChart().getTitle().getTx().getStrRef().setF("Ops!$A$1");
      assertEquals(
          "Fallback title", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      cachedTitleChart.getCTChart().getTitle().getTx().getStrRef().addNewStrCache();
      assertEquals(
          "Fallback title", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      cachedTitleChart
          .getCTChart()
          .getTitle()
          .getTx()
          .getStrRef()
          .getStrCache()
          .addNewPt()
          .setV(" ");
      assertEquals(
          "Fallback title", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      cachedTitleChart
          .getCTChart()
          .getTitle()
          .getTx()
          .getStrRef()
          .getStrCache()
          .getPtArray(0)
          .setV("Cached title");
      assertEquals(
          "Fallback title", ExcelDrawingChartSupport.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      assertEquals("", ExcelDrawingChartSupport.resolvedTitleFormulaText(cachedTitleChart, "Bad["));

      Object noImageDimensions =
          invoke(
              controller,
              "rasterDimensions",
              Object.class,
              "not-an-image".getBytes(StandardCharsets.UTF_8));
      assertNull(invoke(noImageDimensions, "widthPixels", Integer.class));
      assertNull(invoke(noImageDimensions, "heightPixels", Integer.class));
      Object invalidPngDimensions =
          invoke(
              controller,
              "rasterDimensions",
              Object.class,
              Base64.getDecoder()
                  .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVQ="));
      assertNull(invoke(invalidPngDimensions, "widthPixels", Integer.class));
      assertNull(invoke(invalidPngDimensions, "heightPixels", Integer.class));
      assertEquals(
          "Ops!$A$1",
          ExcelDrawingChartSupport.titleSummary(
              new ExcelChartSnapshot.Title.Formula("Ops!$A$1", "")));
      assertEquals(
          "Fallback title",
          ExcelDrawingChartSupport.titleSummary(
              new ExcelChartSnapshot.Title.Formula("Ops!$A$1", "Fallback title")));

      var formulaSeriesTitle =
          org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx.Factory.newInstance();
      formulaSeriesTitle.addNewStrRef().setF("Ops!$A$1");
      assertEquals(
          new ExcelChartSnapshot.Title.Formula("Ops!$A$1", ""),
          ExcelDrawingChartSupport.snapshotSeriesTitle(formulaSeriesTitle));
      formulaSeriesTitle.getStrRef().addNewStrCache();
      assertEquals(
          new ExcelChartSnapshot.Title.Formula("Ops!$A$1", ""),
          ExcelDrawingChartSupport.snapshotSeriesTitle(formulaSeriesTitle));

      org.apache.poi.xssf.usermodel.XSSFChart applyTitleChart =
          drawing.createChart(poiAnchor(drawing, 16, 7, 20, 13));
      var categoryAxis =
          applyTitleChart.createCategoryAxis(
              org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
      var valueAxis =
          applyTitleChart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
      var barData =
          (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData)
              applyTitleChart.createData(
                  org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR, categoryAxis, valueAxis);
      var barSeries =
          (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series)
              barData.addSeries(
                  org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                      new String[] {"Only"}),
                  org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                      new Double[] {1d}));
      Object noTitle =
          ExcelDrawingChartSupport.prepareSeriesTitle(sheet, new ExcelChartDefinition.Title.None());
      ExcelDrawingChartSupport.applySeriesTitle(barSeries, (PreparedSeriesTitle) noTitle);
      assertFalse(barSeries.getCTBarSer().isSetTx());

      org.apache.poi.xssf.usermodel.XSSFChart applyLineTitleChart =
          drawing.createChart(poiAnchor(drawing, 21, 7, 25, 13));
      var lineCategoryAxis =
          applyLineTitleChart.createCategoryAxis(
              org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
      var lineValueAxis =
          applyLineTitleChart.createValueAxis(
              org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
      var lineData =
          (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData)
              applyLineTitleChart.createData(
                  org.apache.poi.xddf.usermodel.chart.ChartTypes.LINE,
                  lineCategoryAxis,
                  lineValueAxis);
      var lineSeries =
          (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series)
              lineData.addSeries(
                  org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                      new String[] {"Only"}),
                  org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                      new Double[] {1d}));
      ExcelDrawingChartSupport.applySeriesTitle(lineSeries, (PreparedSeriesTitle) noTitle);
      assertFalse(lineSeries.getCTLineSer().isSetTx());

      org.apache.poi.xssf.usermodel.XSSFChart applyPieTitleChart =
          drawing.createChart(poiAnchor(drawing, 26, 7, 30, 13));
      var pieData =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData)
              applyPieTitleChart.createData(
                  org.apache.poi.xddf.usermodel.chart.ChartTypes.PIE, null, null);
      var pieSeries =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series)
              pieData.addSeries(
                  org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                      new String[] {"Only"}),
                  org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
                      new Double[] {1d}));
      pieSeries.setTitle("Pie", null);
      ExcelDrawingChartSupport.applySeriesTitle(pieSeries, (PreparedSeriesTitle) noTitle);
      assertFalse(pieSeries.getCTPieSer().isSetTx());

      org.apache.poi.xssf.usermodel.XSSFChart bareBarChart =
          drawing.createChart(poiAnchor(drawing, 20, 0, 24, 6));
      assertFalse(ExcelDrawingChartSupport.barVaryColors(bareBarChart));
      bareBarChart.getCTChart().getPlotArea().addNewBarChart();
      assertFalse(ExcelDrawingChartSupport.barVaryColors(bareBarChart));

      org.apache.poi.xssf.usermodel.XSSFChart bareLineChart =
          drawing.createChart(poiAnchor(drawing, 25, 0, 29, 6));
      assertFalse(ExcelDrawingChartSupport.lineVaryColors(bareLineChart));
      bareLineChart.getCTChart().getPlotArea().addNewLineChart();
      assertFalse(ExcelDrawingChartSupport.lineVaryColors(bareLineChart));

      org.apache.poi.xssf.usermodel.XSSFChart barePieChart =
          drawing.createChart(poiAnchor(drawing, 30, 0, 34, 6));
      assertFalse(ExcelDrawingChartSupport.pieVaryColors(barePieChart));
      barePieChart.getCTChart().getPlotArea().addNewPieChart();
      assertFalse(ExcelDrawingChartSupport.pieVaryColors(barePieChart));
    }
  }
}
