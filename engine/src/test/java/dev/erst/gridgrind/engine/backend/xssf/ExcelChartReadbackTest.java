package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.function.UnaryOperator;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused regression tests for authoritative chart readback behavior. */
class ExcelChartReadbackTest {

  @Test
  void referenceReadbackUsesStoredCachesWhenCurrentCellsDiffer() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-stored-readback-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Summary");
      seedFormulaBackedChartData(sheet);
      workbook.formulas().evaluateAll();
      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.lineChart(
                  "ProjectedLoad",
                  anchor(1, 5, 10, 18),
                  new ResolvedChartTitle.Text("Projected Load"),
                  new ChartLegendSpec.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  List.of(
                      new ResolvedChart.Series(
                          null,
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("C2:C4")))));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    rewriteWorkbookEntry(
        workbookPath,
        "/xl/charts/chart1.xml",
        xml -> xml.replace(">4.0<", ">0.0<").replace(">6.0<", ">0.0<").replace(">10.0<", ">0.0<"));

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ChartReport chart = reopened.sheet("Summary").drawings().charts().getFirst();
      ChartReport.Line line = ExcelChartTestSupport.singlePlot(chart, ChartReport.Line.class);
      ChartReport.DataSource.NumericReference values =
          assertInstanceOf(
              ChartReport.DataSource.NumericReference.class, line.series().getFirst().values());

      assertEquals(Optional.of("C2:C4"), values.formula());
      assertEquals(
          List.of("0.0", "0.0", "0.0"), ChartCacheAssertions.denseNumericValues(values.cache()));
    }
  }

  @Test
  void chartAuthoringPersistsEvaluatedFormulaBackedValueCaches() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-live-authoring-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Summary");
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
              new dev.erst.gridgrind.core.model.value.ResolvedCellText("Owner"));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("B1"),
              new dev.erst.gridgrind.core.model.value.ResolvedCellText("Projected"));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("A2"),
              new dev.erst.gridgrind.core.model.value.ResolvedCellText("Ari"));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
              new dev.erst.gridgrind.core.model.value.ResolvedCellText("Bo"));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("A4"),
              new dev.erst.gridgrind.core.model.value.ResolvedCellText("Cy"));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("B2"),
              new ResolvedFormula("40+2"));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("B3"),
              new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(7d));
      sheet
          .cells()
          .setCell(
              new dev.erst.gridgrind.core.vocabulary.CellAddress("B4"),
              new ResolvedFormula("B3*2"));

      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.lineChart(
                  "ProjectedLoad",
                  anchor(1, 5, 10, 18),
                  new ChartTitleSpec.Formula("B1"),
                  new ChartLegendSpec.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  List.of(
                      new ResolvedChart.Series(
                          null,
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFChart chart = reopened.getSheet("Summary").getDrawingPatriarch().getCharts().getFirst();
      XDDFLineChartData lineData = (XDDFLineChartData) chart.getChartSeries().getFirst();

      assertEquals("42.0", lineData.getSeries(0).getValuesData().getPointAt(0).toString());
      assertEquals("7.0", lineData.getSeries(0).getValuesData().getPointAt(1).toString());
      assertEquals("14.0", lineData.getSeries(0).getValuesData().getPointAt(2).toString());
    }
  }

  private static void seedFormulaBackedChartData(XssfSheet sheet) {
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Owner"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B1"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Hours"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C1"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Projected"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A2"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Ari"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A3"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Bo"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A4"),
            new dev.erst.gridgrind.core.model.value.ResolvedCellText("Cy"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B2"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(2d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B3"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(3d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("B4"),
            new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(5d));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C2"), new ResolvedFormula("B2*2"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C3"), new ResolvedFormula("B3*2"));
    sheet
        .cells()
        .setCell(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("C4"), new ResolvedFormula("B4*2"));
  }

  private static DrawingAnchorSpec.TwoCell anchor(
      int firstRow, int firstColumn, int lastRow, int lastColumn) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(firstRow, firstColumn, 0, 0),
        new DrawingMarkerSpec(lastRow, lastColumn, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
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
