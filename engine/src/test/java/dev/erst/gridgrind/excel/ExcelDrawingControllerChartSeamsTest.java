package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.UnaryOperator;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellBase;
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
              () -> ExcelDrawingChartSupport.requiredDefinedNameFormula(blankName));
      assertTrue(failure.getMessage().contains("Defined name 'BlankSource'"));
    }
  }

  @Test
  void formulaScalarDecoderHandlesBlankAndRejectsMissingCachedResults() {
    assertEquals(
        new ExcelDrawingController.CellScalar(ExcelDrawingController.CellScalarKind.STRING, "", 0d),
        ExcelDrawingChartSupport.scalarFromFormula(
            new FormulaProbeCell(CellType.BLANK, "", 0d, false)));
    assertEquals(
        new ExcelDrawingController.CellScalar(ExcelDrawingController.CellScalarKind.STRING, "", 0d),
        ExcelDrawingChartSupport.scalarFromFormula(
            new FormulaProbeCell(CellType._NONE, "", 0d, false)));

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelDrawingChartSupport.scalarFromFormula(
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

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Charts");
      assertEquals(List.of(), sheet.charts());
      assertEquals(List.of(), sheet.drawingObjects());
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

  /** Minimal formula cell stub for probing scalar-decoding branches without reflection. */
  private static final class FormulaProbeCell extends CellBase {
    private final CellType cachedFormulaResultType;
    private final String stringValue;
    private final double numericValue;
    private final boolean booleanValue;

    private FormulaProbeCell(
        CellType cachedFormulaResultType,
        String stringValue,
        double numericValue,
        boolean booleanValue) {
      this.cachedFormulaResultType = cachedFormulaResultType;
      this.stringValue = stringValue;
      this.numericValue = numericValue;
      this.booleanValue = booleanValue;
    }

    @Override
    public CellType getCellType() {
      return CellType.FORMULA;
    }

    @Override
    public CellType getCachedFormulaResultType() {
      return cachedFormulaResultType;
    }

    @Override
    public int getColumnIndex() {
      return 0;
    }

    @Override
    public int getRowIndex() {
      return 0;
    }

    @Override
    public org.apache.poi.ss.usermodel.Sheet getSheet() {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.Row getRow() {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getCellFormula() {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getStringCellValue() {
      return stringValue;
    }

    @Override
    public double getNumericCellValue() {
      return numericValue;
    }

    @Override
    public boolean getBooleanCellValue() {
      return booleanValue;
    }

    @Override
    public java.util.Date getDateCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public java.time.LocalDateTime getLocalDateTimeCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.RichTextString getRichStringCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public byte getErrorCellValue() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellValue(boolean value) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellErrorValue(byte value) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellStyle(org.apache.poi.ss.usermodel.CellStyle style) {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.CellStyle getCellStyle() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setAsActiveCell() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setCellComment(org.apache.poi.ss.usermodel.Comment comment) {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.Comment getCellComment() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void removeCellComment() {
      throw new UnsupportedOperationException();
    }

    @Override
    public org.apache.poi.ss.usermodel.Hyperlink getHyperlink() {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setHyperlink(org.apache.poi.ss.usermodel.Hyperlink link) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void removeHyperlink() {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellTypeImpl(CellType cellType) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellFormulaImpl(String formula) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void removeFormulaImpl() {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(double value) {
      throw new UnsupportedOperationException();
    }

    @Override
    @SuppressWarnings("PMD.ReplaceJavaUtilDate")
    protected void setCellValueImpl(java.util.Date value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(java.time.LocalDateTime value) {
      throw new UnsupportedOperationException();
    }

    @Override
    @SuppressWarnings("PMD.ReplaceJavaUtilCalendar")
    protected void setCellValueImpl(java.util.Calendar value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(String value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected void setCellValueImpl(org.apache.poi.ss.usermodel.RichTextString value) {
      throw new UnsupportedOperationException();
    }

    @Override
    protected SpreadsheetVersion getSpreadsheetVersion() {
      return SpreadsheetVersion.EXCEL2007;
    }

    @Override
    public org.apache.poi.ss.util.CellRangeAddress getArrayFormulaRange() {
      throw new UnsupportedOperationException();
    }

    @Override
    public boolean isPartOfArrayFormulaGroup() {
      return false;
    }
  }
}
