package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Unit tests for the package-owned POI chart translation seam. */
class ExcelChartPoiBridgeTest {
  @Test
  void convertsAllModeledPoiChartEnumsAndTokens() {
    assertEquals(
        ExcelChartBarDirection.COLUMN, ExcelChartPoiBridge.fromPoiBarDirection(BarDirection.COL));
    assertEquals(
        ExcelChartBarDirection.BAR, ExcelChartPoiBridge.fromPoiBarDirection(BarDirection.BAR));
    assertEquals(
        BarDirection.COL, ExcelChartPoiBridge.toPoiBarDirection(ExcelChartBarDirection.COLUMN));
    assertEquals(
        BarDirection.BAR, ExcelChartPoiBridge.toPoiBarDirection(ExcelChartBarDirection.BAR));

    assertEquals(
        ExcelChartLegendPosition.BOTTOM,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.BOTTOM));
    assertEquals(
        ExcelChartLegendPosition.LEFT,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.LEFT));
    assertEquals(
        ExcelChartLegendPosition.RIGHT,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.RIGHT));
    assertEquals(
        ExcelChartLegendPosition.TOP,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.TOP));
    assertEquals(
        ExcelChartLegendPosition.TOP_RIGHT,
        ExcelChartPoiBridge.fromPoiLegendPosition(LegendPosition.TOP_RIGHT));
    assertEquals(
        LegendPosition.BOTTOM,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.BOTTOM));
    assertEquals(
        LegendPosition.LEFT,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.LEFT));
    assertEquals(
        LegendPosition.RIGHT,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.RIGHT));
    assertEquals(
        LegendPosition.TOP, ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.TOP));
    assertEquals(
        LegendPosition.TOP_RIGHT,
        ExcelChartPoiBridge.toPoiLegendPosition(ExcelChartLegendPosition.TOP_RIGHT));

    assertEquals(
        ExcelChartDisplayBlanksAs.GAP,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(
            org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs.INT_GAP, "gap"));
    assertEquals(
        ExcelChartDisplayBlanksAs.SPAN,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(
            org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs.INT_SPAN, "span"));
    assertEquals(
        ExcelChartDisplayBlanksAs.ZERO,
        ExcelChartPoiBridge.fromPoiDisplayBlanks(
            org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs.INT_ZERO, "zero"));
    IllegalArgumentException unsupportedDisplayBlanks =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelChartPoiBridge.fromPoiDisplayBlanks(99, "bogus"));
    assertTrue(unsupportedDisplayBlanks.getMessage().contains("Unsupported displayBlanksAs token"));
    assertEquals(
        DisplayBlanks.GAP, ExcelChartPoiBridge.toPoiDisplayBlanks(ExcelChartDisplayBlanksAs.GAP));
    assertEquals(
        DisplayBlanks.SPAN, ExcelChartPoiBridge.toPoiDisplayBlanks(ExcelChartDisplayBlanksAs.SPAN));
    assertEquals(
        DisplayBlanks.ZERO, ExcelChartPoiBridge.toPoiDisplayBlanks(ExcelChartDisplayBlanksAs.ZERO));

    assertEquals(
        ExcelChartAxisPosition.BOTTOM,
        ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.BOTTOM));
    assertEquals(
        ExcelChartAxisPosition.LEFT, ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.LEFT));
    assertEquals(
        ExcelChartAxisPosition.RIGHT, ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.RIGHT));
    assertEquals(
        ExcelChartAxisPosition.TOP, ExcelChartPoiBridge.fromPoiAxisPosition(AxisPosition.TOP));

    assertEquals(
        ExcelChartAxisCrosses.AUTO_ZERO,
        ExcelChartPoiBridge.fromPoiAxisCrosses(AxisCrosses.AUTO_ZERO));
    assertEquals(
        ExcelChartAxisCrosses.MAX, ExcelChartPoiBridge.fromPoiAxisCrosses(AxisCrosses.MAX));
    assertEquals(
        ExcelChartAxisCrosses.MIN, ExcelChartPoiBridge.fromPoiAxisCrosses(AxisCrosses.MIN));

    assertEquals("AREA", ExcelChartPoiBridge.canonicalPlotTypeToken("XDDFAreaChartData"));
    assertEquals("CUSTOMPLOT", ExcelChartPoiBridge.canonicalPlotTypeToken("CustomPlot"));
  }

  @Test
  void classifiesConcretePlotFamiliesAndRejectsUnsupportedAxisKinds() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart barChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));
      XDDFCategoryAxis barCategoryAxis = barChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis barValueAxis = barChart.createValueAxis(AxisPosition.LEFT);
      XDDFBarChartData barData =
          (XDDFBarChartData) barChart.createData(ChartTypes.BAR, barCategoryAxis, barValueAxis);

      XSSFChart lineChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 7, 1, 12, 10));
      XDDFCategoryAxis lineCategoryAxis = lineChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis lineValueAxis = lineChart.createValueAxis(AxisPosition.LEFT);
      XDDFLineChartData lineData =
          (XDDFLineChartData)
              lineChart.createData(ChartTypes.LINE, lineCategoryAxis, lineValueAxis);

      XSSFChart pieChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 13, 1, 18, 10));
      XDDFPieChartData pieData = (XDDFPieChartData) pieChart.createData(ChartTypes.PIE, null, null);

      assertEquals("BAR", ExcelChartPoiBridge.plotTypeToken(barData));
      assertEquals("LINE", ExcelChartPoiBridge.plotTypeToken(lineData));
      assertEquals("PIE", ExcelChartPoiBridge.plotTypeToken(pieData));
      assertEquals(ExcelChartAxisKind.CATEGORY, ExcelChartPoiBridge.axisKind(lineCategoryAxis));
      assertEquals(ExcelChartAxisKind.VALUE, ExcelChartPoiBridge.axisKind(lineValueAxis));

      XDDFDateAxis dateAxis = lineChart.createDateAxis(AxisPosition.TOP);
      IllegalArgumentException unsupportedAxis =
          assertThrows(
              IllegalArgumentException.class, () -> ExcelChartPoiBridge.axisKind(dateAxis));
      assertTrue(
          unsupportedAxis
              .getMessage()
              .contains("outside the current modeled simple-chart contract"));
    }
  }

  private static void seedData(XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(18d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
  }
}
