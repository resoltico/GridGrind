package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.Optional;
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
    DrawingAdapter controller = new DrawingAdapter();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      workbook.createSheet("Other");
      sheet.createRow(0).createCell(0).setCellValue("Fallback title");

      assertThrows(
          IllegalArgumentException.class,
          () ->
              ChartAdapter.requiredDefinedNameFormula(
                  new DefinedNameStub("BlankStringSource", " ", -1)));

      org.apache.poi.ss.usermodel.Name workbookScoped = workbook.createName();
      workbookScoped.setNameName("ScopedSource");
      workbookScoped.setRefersToFormula("Ops!$A$1");
      assertEquals(
          Optional.of(workbookScoped),
          ChartAdapter.resolveDefinedNameReference(sheet, "ScopedSource"));
      org.apache.poi.ss.usermodel.Name otherSheetScoped = workbook.createName();
      otherSheetScoped.setNameName("OtherOnly");
      otherSheetScoped.setSheetIndex(workbook.getSheetIndex("Other"));
      otherSheetScoped.setRefersToFormula("Other!$A$1");
      assertEquals(Optional.empty(), ChartAdapter.resolveDefinedNameReference(sheet, "OtherOnly"));
      assertEquals(Optional.empty(), ChartAdapter.resolveDefinedNameReference(sheet, "Bad-1"));

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      org.apache.poi.xssf.usermodel.XSSFChart firstChart =
          drawing.createChart(poiAnchor(drawing, 0, 0, 4, 6));
      firstChart.getGraphicFrame().setName("First");
      org.apache.poi.xssf.usermodel.XSSFChart secondChart =
          drawing.createChart(poiAnchor(drawing, 5, 0, 9, 6));
      secondChart.getGraphicFrame().setName("Second");
      assertSame(
          secondChart, ChartAdapter.chartForGraphicFrame(drawing, secondChart.getGraphicFrame()));

      org.apache.poi.xssf.usermodel.XSSFChart blankTitleChart =
          drawing.createChart(poiAnchor(drawing, 10, 0, 14, 6));
      blankTitleChart.setTitleText(" ");
      assertNotNull(blankTitleChart.getTitleText());
      assertTrue(blankTitleChart.getTitleText().getString().isBlank());
      assertInstanceOf(ChartReport.Title.None.class, ChartAdapter.snapshotTitle(blankTitleChart));
      org.apache.poi.xssf.usermodel.XSSFChart missingTextChart =
          drawing.createChart(poiAnchor(drawing, 15, 0, 19, 6));
      missingTextChart.getCTChart().addNewTitle().addNewTx();
      assertInstanceOf(ChartReport.Title.None.class, ChartAdapter.snapshotTitle(missingTextChart));
      org.apache.poi.xssf.usermodel.XSSFChart textTitleChart =
          drawing.createChart(poiAnchor(drawing, 15, 7, 19, 13));
      textTitleChart.setTitleText("Solid");
      assertEquals(new ChartReport.Title.Text("Solid"), ChartAdapter.snapshotTitle(textTitleChart));

      org.apache.poi.xssf.usermodel.XSSFChart cachedTitleChart =
          drawing.createChart(poiAnchor(drawing, 20, 0, 24, 6));
      assertEquals("", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$99"));
      cachedTitleChart.getCTChart().addNewTitle();
      assertEquals("", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$99"));
      cachedTitleChart.getCTChart().getTitle().addNewTx();
      assertEquals("", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$99"));
      cachedTitleChart.getCTChart().getTitle().getTx().addNewStrRef();
      cachedTitleChart.getCTChart().getTitle().getTx().getStrRef().setF("Ops!$A$1");
      assertEquals("Fallback title", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      cachedTitleChart.getCTChart().getTitle().getTx().getStrRef().addNewStrCache();
      assertEquals("Fallback title", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      cachedTitleChart
          .getCTChart()
          .getTitle()
          .getTx()
          .getStrRef()
          .getStrCache()
          .addNewPt()
          .setV(" ");
      assertEquals("Fallback title", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      cachedTitleChart
          .getCTChart()
          .getTitle()
          .getTx()
          .getStrRef()
          .getStrCache()
          .getPtArray(0)
          .setV("Cached title");
      assertEquals("Fallback title", ChartAdapter.cachedTitleText(cachedTitleChart, "Ops!$A$1"));
      assertEquals("", ChartAdapter.resolvedTitleFormulaText(cachedTitleChart, "Bad["));

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
          "Ops!$A$1", ChartAdapter.titleSummary(new ChartReport.Title.Formula("Ops!$A$1", "")));
      assertEquals(
          "Fallback title",
          ChartAdapter.titleSummary(new ChartReport.Title.Formula("Ops!$A$1", "Fallback title")));

      var formulaSeriesTitle =
          org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx.Factory.newInstance();
      formulaSeriesTitle.addNewStrRef().setF("Ops!$A$1");
      assertEquals(
          new ChartReport.Title.Formula("Ops!$A$1", ""),
          ChartAdapter.snapshotSeriesTitle(formulaSeriesTitle));
      formulaSeriesTitle.getStrRef().addNewStrCache();
      assertEquals(
          new ChartReport.Title.Formula("Ops!$A$1", ""),
          ChartAdapter.snapshotSeriesTitle(formulaSeriesTitle));

      org.apache.poi.xssf.usermodel.XSSFChart bareBarChart =
          drawing.createChart(poiAnchor(drawing, 20, 0, 24, 6));
      assertFalse(ChartAdapter.barVaryColors(bareBarChart));
      bareBarChart.getCTChart().getPlotArea().addNewBarChart();
      assertFalse(ChartAdapter.barVaryColors(bareBarChart));

      org.apache.poi.xssf.usermodel.XSSFChart bareLineChart =
          drawing.createChart(poiAnchor(drawing, 25, 0, 29, 6));
      assertFalse(ChartAdapter.lineVaryColors(bareLineChart));
      bareLineChart.getCTChart().getPlotArea().addNewLineChart();
      assertFalse(ChartAdapter.lineVaryColors(bareLineChart));

      org.apache.poi.xssf.usermodel.XSSFChart barePieChart =
          drawing.createChart(poiAnchor(drawing, 30, 0, 34, 6));
      assertFalse(ChartAdapter.pieVaryColors(barePieChart));
      barePieChart.getCTChart().getPlotArea().addNewPieChart();
      assertFalse(ChartAdapter.pieVaryColors(barePieChart));
    }
  }
}
