package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.IOException;
import java.util.List;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused branch coverage for the extracted chart-axis registry. */
class ExcelChartAxisRegistryTest {
  @Test
  void registryCreatesAndReusesAxesAcrossPlotFamilies() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart categoryChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 8, 12));
      ExcelChartAxisRegistry categoryRegistry = new ExcelChartAxisRegistry(categoryChart);
      ExcelChartDefinition.Axis categoryAxis =
          new ExcelChartDefinition.Axis(
              ExcelChartAxisKind.CATEGORY,
              ExcelChartAxisPosition.BOTTOM,
              ExcelChartAxisCrosses.AUTO_ZERO,
              true);
      ExcelChartDefinition.Axis valueAxis =
          new ExcelChartDefinition.Axis(
              ExcelChartAxisKind.VALUE,
              ExcelChartAxisPosition.LEFT,
              ExcelChartAxisCrosses.MIN,
              false);
      ExcelChartAxisRegistry.CategoryValueAxes first =
          categoryRegistry.categoryValueAxes(List.of(categoryAxis, valueAxis));
      ExcelChartAxisRegistry.CategoryValueAxes second =
          categoryRegistry.categoryValueAxes(List.of(categoryAxis, valueAxis));
      assertSame(first.categoryAxis(), second.categoryAxis());
      assertSame(first.valueAxis(), second.valueAxis());
      assertEquals(AxisPosition.BOTTOM, first.categoryAxis().getPosition());
      assertEquals(AxisPosition.LEFT, first.valueAxis().getPosition());

      XSSFChart datedChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 13, 8, 24));
      ExcelChartAxisRegistry datedRegistry = new ExcelChartAxisRegistry(datedChart);
      ExcelChartAxisRegistry.CategoryValueAxes datedAxes =
          datedRegistry.categoryValueAxes(
              List.of(
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.DATE,
                      ExcelChartAxisPosition.TOP,
                      ExcelChartAxisCrosses.MAX,
                      true),
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.VALUE,
                      ExcelChartAxisPosition.RIGHT,
                      ExcelChartAxisCrosses.AUTO_ZERO,
                      true)));
      assertEquals(AxisPosition.TOP, datedAxes.categoryAxis().getPosition());
      assertEquals(AxisPosition.RIGHT, datedAxes.valueAxis().getPosition());

      XSSFChart scatterChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 9, 1, 16, 12));
      ExcelChartAxisRegistry scatterRegistry = new ExcelChartAxisRegistry(scatterChart);
      ExcelChartAxisRegistry.ScatterAxes scatterAxes =
          scatterRegistry.scatterAxes(
              List.of(
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.VALUE,
                      ExcelChartAxisPosition.BOTTOM,
                      ExcelChartAxisCrosses.MIN,
                      true),
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.VALUE,
                      ExcelChartAxisPosition.LEFT,
                      ExcelChartAxisCrosses.AUTO_ZERO,
                      true)));
      assertEquals(AxisPosition.BOTTOM, scatterAxes.xAxis().getPosition());
      assertEquals(AxisPosition.LEFT, scatterAxes.yAxis().getPosition());

      XSSFChart surfaceChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 17, 1, 24, 12));
      ExcelChartAxisRegistry surfaceRegistry = new ExcelChartAxisRegistry(surfaceChart);
      ExcelChartAxisRegistry.SurfaceAxes surfaceAxes =
          surfaceRegistry.surfaceAxes(
              List.of(
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.DATE,
                      ExcelChartAxisPosition.TOP,
                      ExcelChartAxisCrosses.MAX,
                      true),
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.VALUE,
                      ExcelChartAxisPosition.LEFT,
                      ExcelChartAxisCrosses.AUTO_ZERO,
                      true),
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.SERIES,
                      ExcelChartAxisPosition.RIGHT,
                      ExcelChartAxisCrosses.MIN,
                      false)));
      assertEquals(AxisPosition.TOP, surfaceAxes.categoryAxis().getPosition());
      assertEquals(AxisPosition.RIGHT, surfaceAxes.seriesAxis().getPosition());

      XSSFChart categorySurfaceChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 25, 1, 32, 12));
      ExcelChartAxisRegistry categorySurfaceRegistry =
          new ExcelChartAxisRegistry(categorySurfaceChart);
      ExcelChartAxisRegistry.SurfaceAxes categoricalSurfaceAxes =
          categorySurfaceRegistry.surfaceAxes(
              List.of(
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.CATEGORY,
                      ExcelChartAxisPosition.BOTTOM,
                      ExcelChartAxisCrosses.AUTO_ZERO,
                      true),
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.VALUE,
                      ExcelChartAxisPosition.LEFT,
                      ExcelChartAxisCrosses.MIN,
                      true),
                  new ExcelChartDefinition.Axis(
                      ExcelChartAxisKind.SERIES,
                      ExcelChartAxisPosition.RIGHT,
                      ExcelChartAxisCrosses.MAX,
                      false)));
      assertEquals(AxisPosition.BOTTOM, categoricalSurfaceAxes.categoryAxis().getPosition());
    }
  }

  @Test
  void registryRejectsInvalidAxisCombinations() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart categoryChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 8, 12));
      ExcelChartAxisRegistry categoryRegistry = new ExcelChartAxisRegistry(categoryChart);
      IllegalArgumentException unsupportedSeriesAxis =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  categoryRegistry.categoryValueAxes(
                      List.of(
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.CATEGORY,
                              ExcelChartAxisPosition.BOTTOM,
                              ExcelChartAxisCrosses.AUTO_ZERO,
                              true),
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.SERIES,
                              ExcelChartAxisPosition.RIGHT,
                              ExcelChartAxisCrosses.MAX,
                              true))));
      assertEquals(
          "Series axis is unsupported for this plot family", unsupportedSeriesAxis.getMessage());

      IllegalArgumentException missingValueAxis =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  categoryRegistry.categoryValueAxes(
                      List.of(
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.CATEGORY,
                              ExcelChartAxisPosition.BOTTOM,
                              ExcelChartAxisCrosses.AUTO_ZERO,
                              true))));
      assertEquals(
          "Plot must declare one category/date axis and one value axis",
          missingValueAxis.getMessage());

      IllegalArgumentException missingCategoryAxis =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  categoryRegistry.categoryValueAxes(
                      List.of(
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.VALUE,
                              ExcelChartAxisPosition.LEFT,
                              ExcelChartAxisCrosses.MIN,
                              true))));
      assertEquals(
          "Plot must declare one category/date axis and one value axis",
          missingCategoryAxis.getMessage());

      XSSFChart scatterChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 9, 1, 16, 12));
      ExcelChartAxisRegistry scatterRegistry = new ExcelChartAxisRegistry(scatterChart);
      IllegalArgumentException scatterCountFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  scatterRegistry.scatterAxes(
                      List.of(
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.VALUE,
                              ExcelChartAxisPosition.BOTTOM,
                              ExcelChartAxisCrosses.AUTO_ZERO,
                              true))));
      assertEquals(
          "Scatter plots must declare exactly two value axes", scatterCountFailure.getMessage());

      XSSFChart surfaceChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 17, 1, 24, 12));
      ExcelChartAxisRegistry surfaceRegistry = new ExcelChartAxisRegistry(surfaceChart);
      IllegalArgumentException missingSeriesAxis =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  surfaceRegistry.surfaceAxes(
                      List.of(
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.CATEGORY,
                              ExcelChartAxisPosition.BOTTOM,
                              ExcelChartAxisCrosses.AUTO_ZERO,
                              true),
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.VALUE,
                              ExcelChartAxisPosition.LEFT,
                              ExcelChartAxisCrosses.MIN,
                              true))));
      assertEquals(
          "Surface plots must declare one category/date axis, one value axis, and one series axis",
          missingSeriesAxis.getMessage());

      IllegalArgumentException missingSurfaceCategoryAxis =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  surfaceRegistry.surfaceAxes(
                      List.of(
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.VALUE,
                              ExcelChartAxisPosition.LEFT,
                              ExcelChartAxisCrosses.MIN,
                              true),
                          new ExcelChartDefinition.Axis(
                              ExcelChartAxisKind.SERIES,
                              ExcelChartAxisPosition.RIGHT,
                              ExcelChartAxisCrosses.MAX,
                              false))));
      assertEquals(
          "Surface plots must declare one category/date axis, one value axis, and one series axis",
          missingSurfaceCategoryAxis.getMessage());
    }
  }
}
