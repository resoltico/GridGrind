package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingChartSupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.UnaryOperator;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused tests for chart-source and relation seams that normal authoring APIs cannot express. */
class ExcelDrawingControllerChartSeamsTest {
  @Test
  void blankDefinedNamesAreRejectedWithProductOwnedErrors() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Name blankName = workbook.createName();
      blankName.setNameName("BlankSource");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelChartSourceSupport.requiredDefinedNameFormula(blankName));
      assertTrue(failure.getMessage().contains("Defined name 'BlankSource'"));
    }
  }

  @Test
  void formulaScalarDecoderHandlesBlankAndRejectsMissingCachedResults() {
    assertEquals(
        new ExcelDrawingController.CellScalar(ExcelDrawingController.CellScalarKind.STRING, "", 0d),
        ExcelChartSourceSupport.scalarFromFormula(
            new FormulaProbeCell(CellType.BLANK, "", 0d, false)));
    assertEquals(
        new ExcelDrawingController.CellScalar(ExcelDrawingController.CellScalarKind.STRING, "", 0d),
        ExcelChartSourceSupport.scalarFromFormula(
            new FormulaProbeCell(CellType._NONE, "", 0d, false)));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelChartSourceSupport.scalarFromFormula(
                    new FormulaProbeCell(CellType.FORMULA, "", 0d, false)));
    assertTrue(failure.getMessage().contains("must expose a cached scalar result"));
  }

  @Test
  void pieVaryColorsAndFrameLessChartRelationsAreHandledExplicitly() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart pieChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));
      pieChart.getGraphicFrame().setName("LiteralPie");
      var pieData =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData)
              pieChart.createData(ChartTypes.PIE, null, null);
      pieData.setVaryColors(true);
      pieData.addSeries(
          XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb", "Mar"}),
          XDDFDataSourcesFactory.fromArray(new Double[] {10d, 18d, 15d}));
      pieChart.plot(pieData);

      assertTrue(ExcelDrawingChartSupport.pieVaryColors(pieChart));
      pieChart.getCTChart().getPlotArea().getPieChartArray(0).getVaryColors().setVal(false);
      assertFalse(ExcelDrawingChartSupport.pieVaryColors(pieChart));
    }

    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-frameless-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFChart lineChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));
      lineChart.getGraphicFrame().setName("FrameLess");
      XDDFCategoryAxis categoryAxis = lineChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = lineChart.createValueAxis(AxisPosition.LEFT);
      valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
      var lineData = lineChart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
      lineData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4")));
      lineChart.plot(lineData);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    rewriteWorkbookEntry(
        workbookPath,
        "/xl/drawings/drawing1.xml",
        xml ->
            xml.replaceFirst(
                "(?s)<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr[^>]*name=\"FrameLess\".*?</xdr:graphicFrame>",
                ""));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = workbook.sheet("Charts");
      assertEquals(List.of(), sheet.drawings().charts());
      assertEquals(List.of(), sheet.drawings().drawingObjects());
    }
  }

  @Test
  void liveChartReadbackIgnoresUnrelatedFrameLessChartRelations() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-mixed-frameless-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart liveChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));
      liveChart.getGraphicFrame().setName("LiveChart");
      XDDFCategoryAxis liveCategoryAxis = liveChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis liveValueAxis = liveChart.createValueAxis(AxisPosition.LEFT);
      liveValueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
      var liveData = liveChart.createData(ChartTypes.LINE, liveCategoryAxis, liveValueAxis);
      liveData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4")));
      liveChart.plot(liveData);

      XSSFChart frameLessChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 8, 1, 13, 10));
      frameLessChart.getGraphicFrame().setName("FrameLess");
      XDDFCategoryAxis frameLessCategoryAxis =
          frameLessChart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis frameLessValueAxis = frameLessChart.createValueAxis(AxisPosition.LEFT);
      frameLessValueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
      var frameLessData =
          frameLessChart.createData(ChartTypes.LINE, frameLessCategoryAxis, frameLessValueAxis);
      frameLessData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4")));
      frameLessChart.plot(frameLessData);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    rewriteWorkbookEntry(
        workbookPath,
        "/xl/drawings/drawing1.xml",
        xml ->
            xml.replaceFirst(
                "(?s)<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr[^>]*name=\"FrameLess\".*?</xdr:graphicFrame>",
                ""));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = workbook.sheet("Charts");
      assertEquals(
          List.of("LiveChart"),
          sheet.drawings().charts().stream().map(ExcelChartSnapshot::name).toList());
      assertEquals(
          List.of("LiveChart"),
          sheet.drawings().drawingObjects().stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .toList());
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

  private static void rewriteWorkbookEntry(
      Path workbookPath, String entryPath, UnaryOperator<String> transformer) throws IOException {
    try (var fileSystem = FileSystems.newFileSystem(workbookPath)) {
      Path entry = fileSystem.getPath(entryPath);
      String original = Files.readString(entry);
      String updated = transformer.apply(original);
      assertNotEquals(original, updated);
      Files.writeString(entry, updated);
    }
  }
}
