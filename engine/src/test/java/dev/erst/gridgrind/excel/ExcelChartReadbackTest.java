package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.UnaryOperator;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
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

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Summary");
      seedFormulaBackedChartData(sheet);
      workbook.formulas().evaluateAll();
      sheet.setChart(
          new ExcelChartDefinition.Line(
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
                      new ExcelChartDefinition.DataSource("A2:A4"),
                      new ExcelChartDefinition.DataSource("C2:C4")))));
      workbook.save(workbookPath);
    }

    rewriteWorkbookEntry(
        workbookPath,
        "/xl/charts/chart1.xml",
        xml -> xml.replace(">4.0<", ">0.0<").replace(">6.0<", ">0.0<").replace(">10.0<", ">0.0<"));

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      ExcelChartSnapshot.Line chart =
          assertInstanceOf(
              ExcelChartSnapshot.Line.class, reopened.sheet("Summary").charts().getFirst());
      ExcelChartSnapshot.DataSource.NumericReference values =
          assertInstanceOf(
              ExcelChartSnapshot.DataSource.NumericReference.class,
              chart.series().getFirst().values());

      assertEquals("C2:C4", values.formula());
      assertEquals(List.of("4.0", "6.0", "10.0"), values.cachedValues());
    }
  }

  private static void seedFormulaBackedChartData(ExcelSheet sheet) {
    sheet.setCell("A1", ExcelCellValue.text("Owner"));
    sheet.setCell("B1", ExcelCellValue.text("Hours"));
    sheet.setCell("C1", ExcelCellValue.text("Projected"));
    sheet.setCell("A2", ExcelCellValue.text("Ari"));
    sheet.setCell("A3", ExcelCellValue.text("Bo"));
    sheet.setCell("A4", ExcelCellValue.text("Cy"));
    sheet.setCell("B2", ExcelCellValue.number(2d));
    sheet.setCell("B3", ExcelCellValue.number(3d));
    sheet.setCell("B4", ExcelCellValue.number(5d));
    sheet.setCell("C2", ExcelCellValue.formula("B2*2"));
    sheet.setCell("C3", ExcelCellValue.formula("B3*2"));
    sheet.setCell("C4", ExcelCellValue.formula("B4*2"));
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
