package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import java.io.IOException;
import java.util.List;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Focused coverage for no-runtime chart wrapper overloads and fallback helper seams. */
class ExcelChartWrapperCoverageTest {
  @Test
  void chartMutationAndSnapshotWrappersDelegateWithoutAnExplicitFormulaRuntime()
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartDefinition definition = lineChartDefinition("OpsChart");

      ExcelChartMutationSupport.validateChart(sheet, definition);
      PreparedSeriesTitle preparedTitle =
          ExcelChartMutationSupport.prepareSeriesTitle(
              sheet, new ExcelChartDefinition.Title.Formula("B1"));
      ExcelChartSnapshot.Title.Formula resolvedSeriesTitle =
          assertInstanceOf(
              ExcelChartSnapshot.Title.Formula.class,
              ExcelDrawingChartSupport.snapshotSeriesTitle(
                  sheet,
                  formulaSeriesTitle("Charts!$B$1"),
                  ExcelFormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator())));

      assertInstanceOf(PreparedSeriesTitleFormula.class, preparedTitle);
      assertEquals("Plan", resolvedSeriesTitle.cachedText());

      ExcelChartMutationSupport.createChart(sheet, definition);
      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFChart chart = drawing.getCharts().getFirst();
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();

      assertEquals("OpsChart", ExcelDrawingChartSupport.snapshotChart(chart, graphicFrame).name());
      assertEquals(
          "OpsChart",
          ExcelDrawingChartSupport.snapshotChartDrawingObject(chart, graphicFrame).name());
      assertInstanceOf(
          ExcelDrawingObjectSnapshot.Chart.class,
          ExcelDrawingSnapshotSupport.snapshot(drawing, graphicFrame));
    }
  }

  @Test
  void controllerSheetSupportAndFacadeWrappersCoverDefaultRuntimeDelegation() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet controllerSheet = workbook.createSheet("Controller");
      ExcelChartTestSupport.seedChartData(controllerSheet);
      ExcelDrawingController controller = new ExcelDrawingController();
      controller.setChart(controllerSheet, lineChartDefinition("ControllerChart"));

      ExcelSheetDrawingSupport support = new ExcelSheetDrawingSupport(controllerSheet, controller);
      assertEquals(1, controller.charts(controllerSheet).size());
      assertEquals(1, support.charts().size());

      assertEquals(
          List.of("cached-only"),
          ExcelChartPlotSnapshotSupport.resolvedOrCachedReferenceValues(
              controllerSheet, " ", new ReferenceDataSource(" ", List.of("cached-only"))));

      XSSFSheet facadeSheet = workbook.createSheet("Facade");
      ExcelChartTestSupport.seedChartData(facadeSheet);
      ExcelChartDefinition facadeChart = lineChartDefinition("FacadeChart");
      ExcelDrawingChartSupport.validateChart(facadeSheet, facadeChart);
      ExcelDrawingChartSupport.createChart(facadeSheet, facadeChart);

      assertEquals(1, new ExcelDrawingController().charts(facadeSheet).size());
    }
  }

  private static ExcelChartDefinition lineChartDefinition(String name) {
    return ExcelChartTestSupport.lineChart(
        name,
        ExcelChartTestSupport.anchor(1, 5, 10, 18),
        new ExcelChartDefinition.Title.Text("Ops"),
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Formula("B1"),
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"),
                null,
                null,
                null,
                null)));
  }

  private static CTSerTx formulaSeriesTitle(String formula) {
    CTSerTx title = CTSerTx.Factory.newInstance();
    title.addNewStrRef().setF(formula);
    return title;
  }

  /** Minimal reference-backed data source stub for chart helper fallback coverage. */
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
