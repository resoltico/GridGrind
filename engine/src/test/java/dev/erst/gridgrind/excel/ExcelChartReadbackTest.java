package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.UnaryOperator;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused regression tests for authoritative chart readback behavior. */
class ExcelChartReadbackTest {
  @Test
  void unresolvedReferenceFormulasFallBackToEmbeddedChartCaches() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");

      assertEquals(
          List.of("cached-a", "cached-b"),
          ExcelChartSnapshotSupport.resolvedOrCachedReferenceValues(
              sheet,
              "'Missing'!A1:A2",
              new ReferenceDataSource("'Missing'!A1:A2", List.of("cached-a", "cached-b"))));
    }
  }

  @Test
  void blankReferenceFormulasAlsoFallBackToEmbeddedChartCaches() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");

      assertEquals(
          List.of("cached-only"),
          ExcelChartSnapshotSupport.resolvedOrCachedReferenceValues(
              sheet, " ", new ReferenceDataSource(" ", List.of("cached-only"))));
    }
  }

  @Test
  void nullReferenceFormulasAlsoFallBackToEmbeddedChartCaches() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");

      assertEquals(
          List.of("cached-null"),
          ExcelChartSnapshotSupport.resolvedOrCachedReferenceValues(
              sheet, null, new ReferenceDataSource(null, List.of("cached-null"))));
    }
  }

  @Test
  void referenceReadbackUsesCurrentCellValuesInsteadOfStaleChartCaches() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-live-readback-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Summary");
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
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    rewriteWorkbookEntry(
        workbookPath,
        "/xl/charts/chart1.xml",
        xml -> xml.replace(">4.0<", ">0.0<").replace(">6.0<", ">0.0<").replace(">10.0<", ">0.0<"));

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelChartSnapshot chart = reopened.sheet("Summary").drawings().charts().getFirst();
      ExcelChartSnapshot.Line line =
          ExcelChartTestSupport.singlePlot(chart, ExcelChartSnapshot.Line.class);
      ExcelChartSnapshot.DataSource.NumericReference values =
          assertInstanceOf(
              ExcelChartSnapshot.DataSource.NumericReference.class,
              line.series().getFirst().values());

      assertEquals("C2:C4", values.formula());
      assertEquals(List.of("4.0", "6.0", "10.0"), values.cachedValues());
    }
  }

  @Test
  void chartAuthoringPersistsEvaluatedFormulaBackedValueCaches() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-live-authoring-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Summary");
      sheet.cells().setCell("A1", new CellValue.Text("Owner"));
      sheet.cells().setCell("B1", new CellValue.Text("Projected"));
      sheet.cells().setCell("A2", new CellValue.Text("Ari"));
      sheet.cells().setCell("A3", new CellValue.Text("Bo"));
      sheet.cells().setCell("A4", new CellValue.Text("Cy"));
      sheet.cells().setCell("B2", new ResolvedFormula("40+2"));
      sheet.cells().setCell("B3", new CellValue.NumberValue(7d));
      sheet.cells().setCell("B4", new ResolvedFormula("B3*2"));

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
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
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

  private static void seedFormulaBackedChartData(ExcelSheet sheet) {
    sheet.cells().setCell("A1", new CellValue.Text("Owner"));
    sheet.cells().setCell("B1", new CellValue.Text("Hours"));
    sheet.cells().setCell("C1", new CellValue.Text("Projected"));
    sheet.cells().setCell("A2", new CellValue.Text("Ari"));
    sheet.cells().setCell("A3", new CellValue.Text("Bo"));
    sheet.cells().setCell("A4", new CellValue.Text("Cy"));
    sheet.cells().setCell("B2", new CellValue.NumberValue(2d));
    sheet.cells().setCell("B3", new CellValue.NumberValue(3d));
    sheet.cells().setCell("B4", new CellValue.NumberValue(5d));
    sheet.cells().setCell("C2", new ResolvedFormula("B2*2"));
    sheet.cells().setCell("C3", new ResolvedFormula("B3*2"));
    sheet.cells().setCell("C4", new ResolvedFormula("B4*2"));
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

  /** Minimal reference-backed data source stub for chart snapshot seam coverage. */
  private static final class ReferenceDataSource implements XDDFDataSource<String> {
    private final String formula;
    private final List<String> values;

    private ReferenceDataSource(String formula, List<String> values) {
      this.formula = formula;
      this.values = List.copyOf(values);
    }

    @Override
    public int getPointCount() {
      return values.size();
    }

    @Override
    public String getPointAt(int index) {
      return values.get(index);
    }

    @Override
    public boolean isLiteral() {
      return false;
    }

    @Override
    public boolean isCellRange() {
      return true;
    }

    @Override
    public boolean isReference() {
      return true;
    }

    @Override
    public boolean isNumeric() {
      return false;
    }

    @Override
    public int getColIndex() {
      return 0;
    }

    @Override
    public String getDataRangeReference() {
      return formula;
    }

    @Override
    public String getFormatCode() {
      return null;
    }
  }
}
