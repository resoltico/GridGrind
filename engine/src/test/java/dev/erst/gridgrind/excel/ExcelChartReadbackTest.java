package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
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
                  new ExcelChartDefinition.Title.Text("Projected Load"),
                  new ExcelChartDefinition.Legend.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  List.of(
                      new ExcelChartDefinition.Series(
                          null,
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("C2:C4")))));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
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
      sheet.cells().setCell("A1", ExcelCellValue.text("Owner"));
      sheet.cells().setCell("B1", ExcelCellValue.text("Projected"));
      sheet.cells().setCell("A2", ExcelCellValue.text("Ari"));
      sheet.cells().setCell("A3", ExcelCellValue.text("Bo"));
      sheet.cells().setCell("A4", ExcelCellValue.text("Cy"));
      sheet.cells().setCell("B2", ExcelCellValue.formula("40+2"));
      sheet.cells().setCell("B3", ExcelCellValue.number(7d));
      sheet.cells().setCell("B4", ExcelCellValue.formula("B3*2"));

      sheet
          .drawings()
          .setChart(
              ExcelChartTestSupport.lineChart(
                  "ProjectedLoad",
                  anchor(1, 5, 10, 18),
                  new ExcelChartDefinition.Title.Formula("B1"),
                  new ExcelChartDefinition.Legend.Hidden(),
                  ExcelChartDisplayBlanksAs.GAP,
                  true,
                  false,
                  List.of(
                      new ExcelChartDefinition.Series(
                          null,
                          ExcelChartTestSupport.ref("A2:A4"),
                          ExcelChartTestSupport.ref("B2:B4")))));
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
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
    sheet.cells().setCell("A1", ExcelCellValue.text("Owner"));
    sheet.cells().setCell("B1", ExcelCellValue.text("Hours"));
    sheet.cells().setCell("C1", ExcelCellValue.text("Projected"));
    sheet.cells().setCell("A2", ExcelCellValue.text("Ari"));
    sheet.cells().setCell("A3", ExcelCellValue.text("Bo"));
    sheet.cells().setCell("A4", ExcelCellValue.text("Cy"));
    sheet.cells().setCell("B2", ExcelCellValue.number(2d));
    sheet.cells().setCell("B3", ExcelCellValue.number(3d));
    sheet.cells().setCell("B4", ExcelCellValue.number(5d));
    sheet.cells().setCell("C2", ExcelCellValue.formula("B2*2"));
    sheet.cells().setCell("C3", ExcelCellValue.formula("B3*2"));
    sheet.cells().setCell("C4", ExcelCellValue.formula("B4*2"));
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int firstRow, int firstColumn, int lastRow, int lastColumn) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(firstRow, firstColumn),
        new ExcelDrawingMarker(lastRow, lastColumn),
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
