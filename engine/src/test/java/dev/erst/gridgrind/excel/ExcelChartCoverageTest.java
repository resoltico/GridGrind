package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.DisplayBlanks;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Coverage tests for Phase 6 chart value types, command shapes, and runtime flows. */
class ExcelChartCoverageTest {
  @Test
  void chartValueObjectsAndReadWriteShapesValidateEveryBranch() {
    ExcelDrawingAnchor.TwoCell anchor = anchor(1, 1, 6, 10);
    ExcelChartDefinition.Series definitionSeries =
        new ExcelChartDefinition.Series(
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.DataSource("A2:A4"),
            new ExcelChartDefinition.DataSource("B2:B4"));
    ExcelChartDefinition.Pie pieDefinition =
        new ExcelChartDefinition.Pie(
            "OpsPie",
            anchor,
            new ExcelChartDefinition.Title.Text("Share"),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            45,
            List.of(definitionSeries));
    assertEquals(45, pieDefinition.firstSliceAngle());
    assertInstanceOf(ExcelChartDefinition.Title.None.class, new ExcelChartDefinition.Title.None());
    assertTrue(
        new ExcelChartDefinition.Series(
                    null,
                    new ExcelChartDefinition.DataSource("A2:A4"),
                    new ExcelChartDefinition.DataSource("B2:B4"))
                .title()
            instanceof ExcelChartDefinition.Title.None);
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Pie(
                "OpsPie",
                anchor,
                new ExcelChartDefinition.Title.Text("Share"),
                new ExcelChartDefinition.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                true,
                -1,
                List.of(definitionSeries)));
    assertThrows(IllegalArgumentException.class, () -> new ExcelChartDefinition.Title.Text(" "));
    assertThrows(IllegalArgumentException.class, () -> new ExcelChartDefinition.DataSource(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Line(
                "OpsLine",
                anchor,
                new ExcelChartDefinition.Title.None(),
                new ExcelChartDefinition.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                false,
                List.of()));

    WorkbookCommand.SetChart setChart = new WorkbookCommand.SetChart("Charts", pieDefinition);
    assertEquals("Charts", setChart.sheetName());
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetChart(" ", pieDefinition));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetChart("Charts", null));

    WorkbookReadCommand.GetCharts getCharts = new WorkbookReadCommand.GetCharts("charts", "Charts");
    assertEquals("charts", getCharts.requestId());
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadCommand.GetCharts(" ", "Charts"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadCommand.GetCharts("charts", " "));

    ExcelChartSnapshot.Axis categoryAxis =
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true);
    ExcelChartSnapshot.Axis valueAxis =
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.MIN,
            false);
    ExcelChartSnapshot.Series lineSeries =
        new ExcelChartSnapshot.Series(
            new ExcelChartSnapshot.Title.Formula("Charts!$B$1", "Plan"),
            new ExcelChartSnapshot.DataSource.StringReference(
                "Charts!$A$2:$A$4", List.of("Jan", "Feb")),
            new ExcelChartSnapshot.DataSource.NumericReference(
                "Charts!$B$2:$B$4", "0.0", List.of("10", "18")));
    ExcelChartSnapshot.Line lineSnapshot =
        new ExcelChartSnapshot.Line(
            "OpsLine",
            anchor,
            new ExcelChartSnapshot.Title.None(),
            new ExcelChartSnapshot.Legend.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            true,
            false,
            List.of(categoryAxis, valueAxis),
            List.of(lineSeries));
    ExcelChartSnapshot.Pie pieSnapshot =
        new ExcelChartSnapshot.Pie(
            "OpsPie",
            anchor,
            new ExcelChartSnapshot.Title.Text("Share"),
            new ExcelChartSnapshot.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            true,
            90,
            List.of(
                new ExcelChartSnapshot.Series(
                    new ExcelChartSnapshot.Title.Text("Actual"),
                    new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan", "Feb")),
                    new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("12", "16")))));
    ExcelChartSnapshot.Unsupported unsupportedSnapshot =
        new ExcelChartSnapshot.Unsupported("OpsArea", anchor, List.of("AREA"), "unsupported");
    assertEquals(2, lineSnapshot.axes().size());
    assertEquals(90, pieSnapshot.firstSliceAngle());
    assertEquals(List.of("AREA"), unsupportedSnapshot.plotTypeTokens());

    List<ExcelChartSnapshot> chartSnapshots = new ArrayList<>(List.of(lineSnapshot, pieSnapshot));
    WorkbookReadResult.ChartsResult chartsResult =
        new WorkbookReadResult.ChartsResult("charts", "Charts", chartSnapshots);
    chartSnapshots.clear();
    assertEquals(2, chartsResult.charts().size());
    List<ExcelChartSnapshot> chartsWithNull = new ArrayList<>();
    chartsWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.ChartsResult("charts", "Charts", chartsWithNull));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Line(
                "OpsLine",
                anchor,
                new ExcelChartSnapshot.Title.None(),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.SPAN,
                true,
                false,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Pie(
                "OpsPie",
                anchor,
                new ExcelChartSnapshot.Title.Text("Share"),
                new ExcelChartSnapshot.Legend.Hidden(),
                ExcelChartDisplayBlanksAs.ZERO,
                false,
                true,
                361,
                List.of(
                    new ExcelChartSnapshot.Series(
                        new ExcelChartSnapshot.Title.Text("Actual"),
                        new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                        new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("12"))))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelChartSnapshot.Unsupported("OpsArea", anchor, List.of(" "), "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelChartSnapshot.DataSource.NumericReference("B2:B4", " ", List.of("1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelChartSnapshot.DataSource.NumericLiteral(" ", List.of("1")));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelChartSnapshot.Title.Formula("Charts!$B$1", null));
  }

  @Test
  void lineAndPieChartsExerciseMutationExecutorAndIntrospection() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      seedChartData(sheet);
      seedChartNamedRanges(workbook);

      executor.apply(
          workbook,
          new WorkbookCommand.SetChart(
              "Charts",
              lineChartDefinition(
                  "OpsLine",
                  anchor(4, 1, 10, 16),
                  new ExcelChartDefinition.Title.Text("Line roadmap"),
                  new ExcelChartDefinition.Title.Formula("B1"))));

      WorkbookReadResult.ChartsResult lineRead =
          assertInstanceOf(
              WorkbookReadResult.ChartsResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetCharts("charts", "Charts")));
      ExcelChartSnapshot.Line initialLine =
          assertInstanceOf(
              ExcelChartSnapshot.Line.class,
              lineRead.charts().stream()
                  .filter(snapshot -> "OpsLine".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(new ExcelChartSnapshot.Title.Text("Line roadmap"), initialLine.title());

      sheet.setChart(
          lineChartDefinition(
              "OpsLine",
              anchor(6, 2, 12, 18),
              new ExcelChartDefinition.Title.Text("Line focus"),
              new ExcelChartDefinition.Title.Text("Actual")));
      ExcelChartSnapshot.Line updatedLine =
          assertInstanceOf(
              ExcelChartSnapshot.Line.class,
              sheet.charts().stream()
                  .filter(snapshot -> "OpsLine".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(anchor(6, 2, 12, 18), updatedLine.anchor());
      assertEquals(new ExcelChartSnapshot.Title.Text("Line focus"), updatedLine.title());

      sheet.setChart(
          pieChartDefinition(
              "OpsPie", anchor(13, 1, 19, 12), new ExcelChartDefinition.Title.Text("Share"), 90));
      ExcelChartSnapshot.Pie initialPie =
          assertInstanceOf(
              ExcelChartSnapshot.Pie.class,
              sheet.charts().stream()
                  .filter(snapshot -> "OpsPie".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(90, initialPie.firstSliceAngle());

      sheet.setChart(
          pieChartDefinition(
              "OpsPie",
              anchor(14, 2, 20, 13),
              new ExcelChartDefinition.Title.Text("Updated share"),
              120));
      ExcelChartSnapshot.Pie updatedPie =
          assertInstanceOf(
              ExcelChartSnapshot.Pie.class,
              sheet.charts().stream()
                  .filter(snapshot -> "OpsPie".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(anchor(14, 2, 20, 13), updatedPie.anchor());
      assertEquals(120, updatedPie.firstSliceAngle());

      sheet.setShape(
          new ExcelShapeDefinition(
              "SwapMe",
              ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
              anchor(21, 1, 24, 4),
              "rect",
              "Queue"));
      sheet.setChart(
          lineChartDefinition(
              "SwapMe",
              anchor(21, 1, 27, 8),
              new ExcelChartDefinition.Title.Text("Swap chart"),
              new ExcelChartDefinition.Title.Text("Plan")));
      assertTrue(
          sheet.drawingObjects().stream()
              .anyMatch(
                  snapshot ->
                      snapshot instanceof ExcelDrawingObjectSnapshot.Chart chart
                          && "SwapMe".equals(chart.name())));
      assertTrue(
          sheet.drawingObjects().stream()
              .noneMatch(
                  snapshot ->
                      snapshot instanceof ExcelDrawingObjectSnapshot.Shape shape
                          && "SwapMe".equals(shape.name())));

      IllegalArgumentException invalidChartTitle =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      lineChartDefinition(
                          "BadTitle",
                          anchor(28, 1, 34, 8),
                          new ExcelChartDefinition.Title.Formula("A2:A4"),
                          new ExcelChartDefinition.Title.Text("Plan"))));
      assertTrue(
          invalidChartTitle
              .getMessage()
              .contains("Chart title formula must resolve to a single cell"));
      assertTrue(sheet.charts().stream().noneMatch(snapshot -> "BadTitle".equals(snapshot.name())));

      List<ExcelChartSnapshot> snapshotsBeforeFailure = sheet.charts();
      IllegalArgumentException invalidSeriesTitle =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      lineChartDefinition(
                          "OpsLine",
                          anchor(8, 3, 14, 19),
                          new ExcelChartDefinition.Title.Text("Broken"),
                          new ExcelChartDefinition.Title.Formula("A2:A4"))));
      assertTrue(
          invalidSeriesTitle
              .getMessage()
              .contains("Series title formula must resolve to a single cell"));
      assertEquals(snapshotsBeforeFailure, sheet.charts());
    }
  }

  @Test
  void unsupportedSinglePlotChartsReadBackAsUnsupported() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-unsupported-single-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Charts");
      seedChartData(sheet);

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      var anchor = drawing.createAnchor(0, 0, 0, 0, 4, 1, 11, 16);
      XSSFChart chart = drawing.createChart(anchor);
      chart.getGraphicFrame().setName("AreaOnly");
      XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
      valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
      var categories =
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4"));
      var values =
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4"));
      var areaData = chart.createData(ChartTypes.AREA, categoryAxis, valueAxis);
      areaData.addSeries(categories, values).setTitle("Plan", null);
      chart.plot(areaData);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelChartSnapshot.Unsupported unsupported =
          assertInstanceOf(
              ExcelChartSnapshot.Unsupported.class, workbook.sheet("Charts").charts().getFirst());
      assertEquals("AreaOnly", unsupported.name());
      assertEquals(List.of("AREA"), unsupported.plotTypeTokens());
    }
  }

  @Test
  void directPoiChartReadbackCoversLiteralSourcesTitlesAndAxisVariants() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-literal-readback-");
    writeLiteralChartReadbackWorkbook(workbookPath);
    assertLiteralChartReadback(workbookPath);
  }

  private static void writeLiteralChartReadbackWorkbook(Path workbookPath) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart cachedFormulaLine =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 8, 12));
      cachedFormulaLine.getGraphicFrame().setName("");
      var cachedLineTitle = cachedFormulaLine.getCTChart().addNewTitle();
      var cachedLineTitleRef = cachedLineTitle.addNewTx().addNewStrRef();
      cachedLineTitleRef.setF("Charts!$B$1");
      var cachedLineTitleCache = cachedLineTitleRef.addNewStrCache();
      cachedLineTitleCache.addNewPtCount().setVal(1);
      var cachedLineTitlePoint = cachedLineTitleCache.addNewPt();
      cachedLineTitlePoint.setIdx(0);
      cachedLineTitlePoint.setV("Plan");
      cachedFormulaLine
          .getOrAddLegend()
          .setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.LEFT);
      cachedFormulaLine.displayBlanksAs(DisplayBlanks.SPAN);
      XDDFCategoryAxis cachedLineCategoryAxis =
          cachedFormulaLine.createCategoryAxis(AxisPosition.TOP);
      XDDFValueAxis cachedLineValueAxis = cachedFormulaLine.createValueAxis(AxisPosition.RIGHT);
      cachedLineValueAxis.setCrosses(AxisCrosses.MAX);
      var cachedLineData =
          cachedFormulaLine.createData(
              ChartTypes.LINE, cachedLineCategoryAxis, cachedLineValueAxis);
      CTSerTx cachedLineSeriesTitle =
          ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series)
                  cachedLineData.addSeries(
                      XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Mar", null}),
                      XDDFDataSourcesFactory.fromArray(new Double[] {10d, 18d, 15d})))
              .getCTLineSer()
              .addNewTx();
      cachedLineSeriesTitle.addNewStrRef().setF("Charts!$C$1");
      cachedFormulaLine.plot(cachedLineData);

      XSSFChart uncachedFormulaLine =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 9, 1, 16, 12));
      uncachedFormulaLine.getGraphicFrame().setName("FormulaNoCache");
      var uncachedLineTitle = uncachedFormulaLine.getCTChart().addNewTitle();
      uncachedLineTitle.addNewTx().addNewStrRef().setF("Charts!$C$1");
      uncachedFormulaLine.displayBlanksAs(DisplayBlanks.ZERO);
      XDDFCategoryAxis uncachedLineCategoryAxis =
          uncachedFormulaLine.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis uncachedLineValueAxis = uncachedFormulaLine.createValueAxis(AxisPosition.LEFT);
      uncachedLineValueAxis.setCrosses(AxisCrosses.MIN);
      var uncachedLineData =
          uncachedFormulaLine.createData(
              ChartTypes.LINE, uncachedLineCategoryAxis, uncachedLineValueAxis);
      uncachedLineData.addSeries(
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4")),
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4")));
      uncachedFormulaLine.plot(uncachedLineData);

      XSSFChart literalPie = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 17, 1, 24, 12));
      literalPie.getGraphicFrame().setName("LiteralPie");
      literalPie.getCTChart().addNewTitle();
      literalPie
          .getOrAddLegend()
          .setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.TOP);
      var literalPieData =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData)
              literalPie.createData(ChartTypes.PIE, null, null);
      literalPieData.setVaryColors(true);
      literalPieData.setFirstSliceAngle(75);
      literalPieData.addSeries(
          XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb", "Mar"}),
          XDDFDataSourcesFactory.fromArray(new Double[] {12d, 16d, 21d}));
      literalPie.plot(literalPieData);

      XSSFChart bottomLegendBar =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 25, 1, 32, 12));
      bottomLegendBar.getGraphicFrame().setName("BottomLegend");
      bottomLegendBar
          .getOrAddLegend()
          .setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);
      XDDFCategoryAxis bottomBarCategoryAxis =
          bottomLegendBar.createCategoryAxis(AxisPosition.BOTTOM);
      XDDFValueAxis bottomBarValueAxis = bottomLegendBar.createValueAxis(AxisPosition.LEFT);
      bottomBarValueAxis.setCrosses(AxisCrosses.MIN);
      var bottomBarData =
          bottomLegendBar.createData(ChartTypes.BAR, bottomBarCategoryAxis, bottomBarValueAxis);
      var bottomBarSeries =
          (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series)
              bottomBarData.addSeries(
                  XDDFDataSourcesFactory.fromStringCellRange(
                      sheet, CellRangeAddress.valueOf("A2:A4")),
                  XDDFDataSourcesFactory.fromNumericCellRange(
                      sheet, CellRangeAddress.valueOf("B2:B4")));
      bottomBarSeries.getCTBarSer().addNewTx();
      bottomLegendBar.plot(bottomBarData);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }
  }

  private static void assertLiteralChartReadback(Path workbookPath) throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Charts");

      ExcelChartSnapshot.Line cachedFormulaLine =
          sheet.charts().stream()
              .filter(snapshot -> snapshot instanceof ExcelChartSnapshot.Line)
              .map(ExcelChartSnapshot.Line.class::cast)
              .filter(snapshot -> snapshot.name().startsWith("Chart-"))
              .findFirst()
              .orElseThrow();
      ExcelChartSnapshot.Title.Formula cachedLineTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, cachedFormulaLine.title());
      assertEquals("Charts!$B$1", cachedLineTitle.formula());
      assertEquals("Plan", cachedLineTitle.cachedText());
      assertEquals(
          ExcelChartLegendPosition.LEFT,
          assertInstanceOf(ExcelChartSnapshot.Legend.Visible.class, cachedFormulaLine.legend())
              .position());
      assertEquals(ExcelChartDisplayBlanksAs.SPAN, cachedFormulaLine.displayBlanksAs());
      assertEquals(
          List.of(ExcelChartAxisPosition.TOP, ExcelChartAxisPosition.RIGHT),
          cachedFormulaLine.axes().stream().map(ExcelChartSnapshot.Axis::position).toList());
      assertEquals(
          List.of(ExcelChartAxisCrosses.AUTO_ZERO, ExcelChartAxisCrosses.MAX),
          cachedFormulaLine.axes().stream().map(ExcelChartSnapshot.Axis::crosses).toList());
      ExcelChartSnapshot.Series cachedLineSeries = cachedFormulaLine.series().getFirst();
      ExcelChartSnapshot.Title.Formula cachedLineSeriesTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, cachedLineSeries.title());
      assertEquals("", cachedLineSeriesTitle.cachedText());
      assertEquals(
          List.of("Jan", "Mar", ""),
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.StringLiteral.class, cachedLineSeries.categories())
              .values());
      assertEquals(
          List.of("10.0", "18.0", "15.0"),
          assertInstanceOf(
                  ExcelChartSnapshot.DataSource.NumericLiteral.class, cachedLineSeries.values())
              .values());

      ExcelChartSnapshot.Line uncachedFormulaLine =
          findChart(sheet.charts(), "FormulaNoCache", ExcelChartSnapshot.Line.class);
      ExcelChartSnapshot.Title.Formula uncachedLineTitle =
          assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, uncachedFormulaLine.title());
      assertEquals("Charts!$C$1", uncachedLineTitle.formula());
      assertEquals("", uncachedLineTitle.cachedText());

      ExcelChartSnapshot.Pie literalPie =
          findChart(sheet.charts(), "LiteralPie", ExcelChartSnapshot.Pie.class);
      assertInstanceOf(ExcelChartSnapshot.Title.None.class, literalPie.title());
      assertEquals(
          ExcelChartLegendPosition.TOP,
          assertInstanceOf(ExcelChartSnapshot.Legend.Visible.class, literalPie.legend())
              .position());
      assertTrue(literalPie.varyColors());
      assertEquals(75, literalPie.firstSliceAngle());

      ExcelChartSnapshot.Bar bottomLegendBar =
          findChart(sheet.charts(), "BottomLegend", ExcelChartSnapshot.Bar.class);
      assertEquals(
          ExcelChartLegendPosition.BOTTOM,
          assertInstanceOf(ExcelChartSnapshot.Legend.Visible.class, bottomLegendBar.legend())
              .position());
      assertTrue(
          bottomLegendBar.series().getFirst().title() instanceof ExcelChartSnapshot.Title.None);
      assertEquals(ExcelChartAxisCrosses.MIN, bottomLegendBar.axes().get(1).crosses());

      ExcelDrawingObjectSnapshot.Chart cachedFormulaDrawingObject =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> snapshot.name().equals(cachedFormulaLine.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals("Plan", cachedFormulaDrawingObject.title());
      ExcelDrawingObjectSnapshot.Chart uncachedFormulaDrawingObject =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> "FormulaNoCache".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals("Charts!$C$1", uncachedFormulaDrawingObject.title());
      ExcelDrawingObjectSnapshot.Chart pieDrawingObject =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Chart.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> "LiteralPie".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals("", pieDrawingObject.title());
    }
  }

  @Test
  void chartSourceResolutionAndScalarBranchesStayDeterministic() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      seedChartData(sheet);
      seedChartNamedRanges(workbook);
      seedChartSourceResolutionFixtures(workbook, sheet);
      assertChartSourceResolutionReadbacks(sheet);
      assertChartSourceResolutionFailures(sheet);
    }
  }

  @Test
  void existingLineAndPieChartsRejectTypeChangesWithoutMutatingState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      seedChartData(sheet);
      seedChartNamedRanges(workbook);

      sheet.setChart(
          lineChartDefinition(
              "OpsLine",
              anchor(4, 1, 10, 16),
              new ExcelChartDefinition.Title.Text("Line roadmap"),
              new ExcelChartDefinition.Title.Text("Plan")));
      List<ExcelChartSnapshot> lineBeforeFailure = sheet.charts();
      IllegalArgumentException lineTypeChange =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      pieChartDefinition(
                          "OpsLine",
                          anchor(4, 1, 10, 16),
                          new ExcelChartDefinition.Title.Text("Type change"),
                          45)));
      assertTrue(lineTypeChange.getMessage().contains("Changing chart type"));
      assertEquals(lineBeforeFailure, sheet.charts());

      sheet.setChart(
          pieChartDefinition(
              "OpsPie", anchor(12, 1, 18, 12), new ExcelChartDefinition.Title.Text("Share"), 90));
      List<ExcelChartSnapshot> pieBeforeFailure = sheet.charts();
      IllegalArgumentException pieTypeChange =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  sheet.setChart(
                      lineChartDefinition(
                          "OpsPie",
                          anchor(12, 1, 18, 12),
                          new ExcelChartDefinition.Title.Text("Type change"),
                          new ExcelChartDefinition.Title.Text("Plan"))));
      assertTrue(pieTypeChange.getMessage().contains("Changing chart type"));
      assertEquals(pieBeforeFailure, sheet.charts());
    }
  }

  @Test
  void chartSeriesTitleMutationAndSeriesRemovalStayDeterministic() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Charts");
      seedChartData(sheet);
      seedChartNamedRanges(workbook);

      sheet.setChart(
          new ExcelChartDefinition.Bar(
              "OpsBarSeries",
              anchor(1, 32, 7, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              false,
              ExcelChartBarDirection.COLUMN,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Text("Plan"),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartPlan")),
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Formula("C1"),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartActual")))));
      sheet.setChart(
          new ExcelChartDefinition.Bar(
              "OpsBarSeries",
              anchor(1, 32, 7, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              false,
              ExcelChartBarDirection.COLUMN,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.None(),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartPlan")))));
      ExcelChartSnapshot.Bar barSeriesChart =
          findChart(sheet.charts(), "OpsBarSeries", ExcelChartSnapshot.Bar.class);
      assertEquals(1, barSeriesChart.series().size());
      assertTrue(
          barSeriesChart.series().getFirst().title() instanceof ExcelChartSnapshot.Title.None);

      sheet.setChart(
          new ExcelChartDefinition.Line(
              "OpsLineSeries",
              anchor(8, 32, 14, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              false,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Formula("B2"),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartPlan")),
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Formula("H30"),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartActual")))));
      ExcelChartSnapshot.Line initialLineSeriesChart =
          findChart(sheet.charts(), "OpsLineSeries", ExcelChartSnapshot.Line.class);
      assertEquals(
          "10.0",
          assertInstanceOf(
                  ExcelChartSnapshot.Title.Formula.class,
                  initialLineSeriesChart.series().getFirst().title())
              .cachedText());
      assertEquals(
          "",
          assertInstanceOf(
                  ExcelChartSnapshot.Title.Formula.class,
                  initialLineSeriesChart.series().get(1).title())
              .cachedText());
      sheet.setChart(
          new ExcelChartDefinition.Line(
              "OpsLineSeries",
              anchor(8, 32, 14, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              false,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.None(),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartPlan")))));
      ExcelChartSnapshot.Line lineSeriesChart =
          findChart(sheet.charts(), "OpsLineSeries", ExcelChartSnapshot.Line.class);
      assertEquals(1, lineSeriesChart.series().size());
      assertTrue(
          lineSeriesChart.series().getFirst().title() instanceof ExcelChartSnapshot.Title.None);

      sheet.setChart(
          new ExcelChartDefinition.Pie(
              "OpsPieSeries",
              anchor(15, 32, 21, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              true,
              45,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Formula("B1"),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartPlan")),
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.Text("Actual"),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartActual")))));
      sheet.setChart(
          new ExcelChartDefinition.Pie(
              "OpsPieSeries",
              anchor(15, 32, 21, 46),
              new ExcelChartDefinition.Title.None(),
              new ExcelChartDefinition.Legend.Hidden(),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              true,
              45,
              List.of(
                  new ExcelChartDefinition.Series(
                      new ExcelChartDefinition.Title.None(),
                      new ExcelChartDefinition.DataSource("ChartCategories"),
                      new ExcelChartDefinition.DataSource("ChartPlan")))));
      ExcelChartSnapshot.Pie pieSeriesChart =
          findChart(sheet.charts(), "OpsPieSeries", ExcelChartSnapshot.Pie.class);
      assertEquals(1, pieSeriesChart.series().size());
      assertTrue(
          pieSeriesChart.series().getFirst().title() instanceof ExcelChartSnapshot.Title.None);
      assertTrue(pieSeriesChart.varyColors());
    }
  }

  @Test
  void malformedAndOrphanedChartFramesDegradeIntoTruthfulReadOnlyFacts() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-graphic-frame-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart orphanChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));
      orphanChart.getGraphicFrame().setName("OrphanFrame");
      var orphanCategories =
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4"));
      var orphanValues =
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4"));
      var orphanData =
          orphanChart.createData(
              ChartTypes.LINE,
              orphanChart.createCategoryAxis(AxisPosition.BOTTOM),
              orphanChart.createValueAxis(AxisPosition.LEFT));
      orphanData.addSeries(orphanCategories, orphanValues).setTitle("Plan", null);
      orphanChart.plot(orphanData);

      XSSFChart brokenChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 7, 1, 12, 10));
      brokenChart.getGraphicFrame().setName("BrokenChart");
      var brokenCategories =
          XDDFDataSourcesFactory.fromStringCellRange(sheet, CellRangeAddress.valueOf("A2:A4"));
      var brokenValues =
          XDDFDataSourcesFactory.fromNumericCellRange(sheet, CellRangeAddress.valueOf("B2:B4"));
      var brokenData =
          brokenChart.createData(
              ChartTypes.LINE,
              brokenChart.createCategoryAxis(AxisPosition.BOTTOM),
              brokenChart.createValueAxis(AxisPosition.LEFT));
      brokenData.addSeries(brokenCategories, brokenValues).setTitle("Plan", null);
      brokenChart.plot(brokenData);
      ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) brokenData.getSeries(0))
          .getCTLineSer()
          .unsetVal();

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    removeFirstChartRelationship(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Charts");

      ExcelDrawingObjectSnapshot.Shape orphanFrame =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Shape.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> "OrphanFrame".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(ExcelDrawingShapeKind.GRAPHIC_FRAME, orphanFrame.kind());

      IllegalArgumentException anchorFailure =
          assertThrows(
              IllegalArgumentException.class,
              () -> sheet.setDrawingObjectAnchor("OrphanFrame", anchor(2, 2, 7, 11)));
      assertTrue(anchorFailure.getMessage().contains("is read-only until a later parity phase"));

      IllegalArgumentException deleteFailure =
          assertThrows(
              IllegalArgumentException.class, () -> sheet.deleteDrawingObject("OrphanFrame"));
      assertTrue(deleteFailure.getMessage().contains("is read-only until a later parity phase"));

      ExcelChartSnapshot.Unsupported brokenChart =
          findChart(sheet.charts(), "BrokenChart", ExcelChartSnapshot.Unsupported.class);
      assertEquals(List.of("LINE"), brokenChart.plotTypeTokens());
      assertTrue(brokenChart.detail().contains("missing its data source"));
    }
  }

  @Test
  void chartDeletionFailureIsVerifiableThroughTheRelationRemovalSeam() throws IOException {
    ExcelDrawingController controller = new ExcelDrawingController((parent, child) -> false);

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));
      chart.getGraphicFrame().setName("OpsChart");

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class, () -> controller.deleteDrawingObject(sheet, "OpsChart"));
      assertTrue(failure.getMessage().contains("Failed to remove chart relation for 'OpsChart'"));
    }
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static void seedChartSourceResolutionFixtures(ExcelWorkbook workbook, ExcelSheet sheet) {
    sheet.setCell("D2", ExcelCellValue.bool(true));
    sheet.setCell("D3", ExcelCellValue.blank());
    workbook
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "NumericCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "B2:B4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "SparseCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "D2:D4")));

    var localCategories = workbook.xssfWorkbook().createName();
    localCategories.setNameName("LocalCategories");
    localCategories.setSheetIndex(workbook.xssfWorkbook().getSheetIndex("Charts"));
    localCategories.setRefersToFormula("A2:A4");

    var multiAreaSource = workbook.xssfWorkbook().createName();
    multiAreaSource.setNameName("MultiAreaSource");
    multiAreaSource.setRefersToFormula("A2:A4,C2:C4");

    var xssfSheet = sheet.xssfSheet();
    xssfSheet.getRow(1).createCell(4).setCellFormula("\"Alpha\"");
    xssfSheet.getRow(2).createCell(4).setCellFormula("1+1");
    xssfSheet.getRow(3).createCell(4).setCellFormula("TRUE");
    xssfSheet.createRow(4).createCell(4).setCellFormula("\"\"");
    xssfSheet.getRow(4).createCell(2).setCellValue(25d);
    workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();

    workbook
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "FormulaCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "E2:E5")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "FormulaValues",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "C2:C5")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ErrorFormulaCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "F2:F4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ErrorCellCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "G2:G4")));

    xssfSheet.getRow(1).createCell(5).setCellFormula("1/0");
    xssfSheet.getRow(2).createCell(5).setCellFormula("1/0");
    xssfSheet.getRow(3).createCell(5).setCellFormula("1/0");
    workbook.xssfWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();
    xssfSheet.getRow(1).createCell(6).setCellErrorValue(FormulaError.NA.getCode());
    xssfSheet.getRow(2).createCell(6).setCellErrorValue(FormulaError.NA.getCode());
    xssfSheet.getRow(3).createCell(6).setCellErrorValue(FormulaError.NA.getCode());

    var implicitSheetSource = workbook.xssfWorkbook().createName();
    implicitSheetSource.setNameName("ImplicitSheetSource");
    implicitSheetSource.setRefersToFormula("H20:H22");

    ExcelSheet sourceSheet = workbook.getOrCreateSheet("Other");
    sourceSheet.setCell("A2", ExcelCellValue.text("North"));
    sourceSheet.setCell("A3", ExcelCellValue.text("South"));
    sourceSheet.setCell("A4", ExcelCellValue.text("West"));
  }

  private static void assertChartSourceResolutionReadbacks(ExcelSheet sheet) {
    sheet.setChart(
        new ExcelChartDefinition.Line(
            "FormulaTitle",
            anchor(4, 16, 10, 30),
            new ExcelChartDefinition.Title.Formula("=B1"),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(
                new ExcelChartDefinition.Series(
                    null,
                    new ExcelChartDefinition.DataSource("NumericCategories"),
                    new ExcelChartDefinition.DataSource("ChartActual")))));
    ExcelChartSnapshot.Line formulaTitleChart =
        findChart(sheet.charts(), "FormulaTitle", ExcelChartSnapshot.Line.class);
    ExcelChartSnapshot.Title.Formula chartTitle =
        assertInstanceOf(ExcelChartSnapshot.Title.Formula.class, formulaTitleChart.title());
    assertEquals("", chartTitle.cachedText());
    assertTrue(
        formulaTitleChart.series().getFirst().title() instanceof ExcelChartSnapshot.Title.None);
    assertInstanceOf(
        ExcelChartSnapshot.DataSource.NumericReference.class,
        formulaTitleChart.series().getFirst().categories());

    sheet.setChart(
        new ExcelChartDefinition.Line(
            "SparseCategoriesChart",
            anchor(11, 16, 17, 30),
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(
                new ExcelChartDefinition.Series(
                    null,
                    new ExcelChartDefinition.DataSource("SparseCategories"),
                    new ExcelChartDefinition.DataSource("ChartActual")))));
    ExcelChartSnapshot.Line sparseChart =
        findChart(sheet.charts(), "SparseCategoriesChart", ExcelChartSnapshot.Line.class);
    assertEquals(
        List.of("true", "", ""),
        ((ExcelChartSnapshot.DataSource.StringReference)
                sparseChart.series().getFirst().categories())
            .cachedValues());

    sheet.setChart(
        new ExcelChartDefinition.Line(
            "FormulaCategoriesChart",
            anchor(18, 16, 24, 30),
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(
                new ExcelChartDefinition.Series(
                    null,
                    new ExcelChartDefinition.DataSource("FormulaCategories"),
                    new ExcelChartDefinition.DataSource("FormulaValues")))));
    ExcelChartSnapshot.Line formulaCategoriesChart =
        findChart(sheet.charts(), "FormulaCategoriesChart", ExcelChartSnapshot.Line.class);
    assertEquals(
        List.of("Alpha", "2.0", "true", ""),
        ((ExcelChartSnapshot.DataSource.StringReference)
                formulaCategoriesChart.series().getFirst().categories())
            .cachedValues());

    assertDoesNotThrow(
        () ->
            sheet.setChart(
                new ExcelChartDefinition.Line(
                    "LocalCategoriesChart",
                    anchor(25, 16, 31, 30),
                    new ExcelChartDefinition.Title.None(),
                    new ExcelChartDefinition.Legend.Hidden(),
                    ExcelChartDisplayBlanksAs.GAP,
                    true,
                    false,
                    List.of(
                        new ExcelChartDefinition.Series(
                            null,
                            new ExcelChartDefinition.DataSource("LocalCategories"),
                            new ExcelChartDefinition.DataSource("ChartActual"))))));

    sheet.setChart(
        new ExcelChartDefinition.Line(
            "ImplicitSheetSourceChart",
            anchor(26, 1, 32, 14),
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(
                new ExcelChartDefinition.Series(
                    null,
                    new ExcelChartDefinition.DataSource("ImplicitSheetSource"),
                    new ExcelChartDefinition.DataSource("ChartActual")))));
    ExcelChartSnapshot.Line implicitSheetSourceChart =
        findChart(sheet.charts(), "ImplicitSheetSourceChart", ExcelChartSnapshot.Line.class);
    assertEquals(
        List.of("", "", ""),
        ((ExcelChartSnapshot.DataSource.StringReference)
                implicitSheetSourceChart.series().getFirst().categories())
            .cachedValues());

    sheet.setChart(
        new ExcelChartDefinition.Line(
            "CrossSheetCategoriesChart",
            anchor(33, 1, 39, 14),
            new ExcelChartDefinition.Title.None(),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            List.of(
                new ExcelChartDefinition.Series(
                    null,
                    new ExcelChartDefinition.DataSource("'Other'!A2:A4"),
                    new ExcelChartDefinition.DataSource("ChartActual")))));
    ExcelChartSnapshot.Line crossSheetCategoriesChart =
        findChart(sheet.charts(), "CrossSheetCategoriesChart", ExcelChartSnapshot.Line.class);
    assertEquals(
        List.of("North", "South", "West"),
        ((ExcelChartSnapshot.DataSource.StringReference)
                crossSheetCategoriesChart.series().getFirst().categories())
            .cachedValues());
  }

  private static void assertChartSourceResolutionFailures(ExcelSheet sheet) {
    IllegalArgumentException stringValues =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "StringValues",
                        anchor(32, 16, 38, 30),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("ChartCategories"),
                                new ExcelChartDefinition.DataSource("LocalCategories"))))));
    assertTrue(
        stringValues.getMessage().contains("Chart value source must resolve to numeric cells"));

    IllegalArgumentException nonContiguous =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "NonContiguous",
                        anchor(39, 16, 45, 30),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("A2:A4,C2:C4"),
                                new ExcelChartDefinition.DataSource("ChartActual"))))));
    assertTrue(nonContiguous.getMessage().contains("one contiguous area"));

    IllegalArgumentException invalidCellReference =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "InvalidCellReference",
                        anchor(47, 1, 53, 14),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("A2:???"),
                                new ExcelChartDefinition.DataSource("ChartActual"))))));
    assertTrue(invalidCellReference.getMessage().contains("one contiguous area"));

    IllegalArgumentException multiAreaDefinedName =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "MultiAreaDefinedName",
                        anchor(53, 16, 59, 30),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("MultiAreaSource"),
                                new ExcelChartDefinition.DataSource("ChartActual"))))));
    assertTrue(multiAreaDefinedName.getMessage().contains("must resolve to one contiguous area"));

    IllegalArgumentException missingSheet =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "MissingSheet",
                        anchor(60, 16, 66, 30),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("'Missing'!A2:A4"),
                                new ExcelChartDefinition.DataSource("ChartActual"))))));
    assertTrue(missingSheet.getMessage().contains("resolves to missing sheet"));

    IllegalArgumentException errorFormulaCategories =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "ErrorFormulaCategoriesChart",
                        anchor(67, 16, 73, 30),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("ErrorFormulaCategories"),
                                new ExcelChartDefinition.DataSource("ChartActual"))))));
    assertTrue(errorFormulaCategories.getMessage().contains("must not cache error values"));

    IllegalArgumentException errorCellCategories =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                sheet.setChart(
                    new ExcelChartDefinition.Line(
                        "ErrorCellCategoriesChart",
                        anchor(74, 16, 80, 30),
                        new ExcelChartDefinition.Title.None(),
                        new ExcelChartDefinition.Legend.Hidden(),
                        ExcelChartDisplayBlanksAs.GAP,
                        true,
                        false,
                        List.of(
                            new ExcelChartDefinition.Series(
                                null,
                                new ExcelChartDefinition.DataSource("ErrorCellCategories"),
                                new ExcelChartDefinition.DataSource("ChartActual"))))));
    assertTrue(errorCellCategories.getMessage().contains("must not contain error values"));
  }

  private static ExcelChartDefinition.Line lineChartDefinition(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      ExcelChartDefinition.Title seriesTitle) {
    return new ExcelChartDefinition.Line(
        name,
        anchor,
        title,
        new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        List.of(
            new ExcelChartDefinition.Series(
                seriesTitle,
                new ExcelChartDefinition.DataSource("ChartCategories"),
                new ExcelChartDefinition.DataSource("ChartPlan"))));
  }

  private static ExcelChartDefinition.Pie pieChartDefinition(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      int firstSliceAngle) {
    return new ExcelChartDefinition.Pie(
        name,
        anchor,
        title,
        new ExcelChartDefinition.Legend.Hidden(),
        ExcelChartDisplayBlanksAs.ZERO,
        false,
        true,
        firstSliceAngle,
        List.of(
            new ExcelChartDefinition.Series(
                new ExcelChartDefinition.Title.Text("Actual"),
                new ExcelChartDefinition.DataSource("ChartCategories"),
                new ExcelChartDefinition.DataSource("ChartActual"))));
  }

  private static void seedChartNamedRanges(ExcelWorkbook workbook) {
    workbook
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "A2:A4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartPlan",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "B2:B4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartActual",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget("Charts", "C2:C4")));
  }

  private static void seedChartData(ExcelSheet sheet) {
    sheet.setCell("A1", ExcelCellValue.text("Month"));
    sheet.setCell("B1", ExcelCellValue.text("Plan"));
    sheet.setCell("C1", ExcelCellValue.text("Actual"));
    sheet.setCell("A2", ExcelCellValue.text("Jan"));
    sheet.setCell("B2", ExcelCellValue.number(10d));
    sheet.setCell("C2", ExcelCellValue.number(12d));
    sheet.setCell("A3", ExcelCellValue.text("Feb"));
    sheet.setCell("B3", ExcelCellValue.number(18d));
    sheet.setCell("C3", ExcelCellValue.number(16d));
    sheet.setCell("A4", ExcelCellValue.text("Mar"));
    sheet.setCell("B4", ExcelCellValue.number(15d));
    sheet.setCell("C4", ExcelCellValue.number(21d));
  }

  private static void seedChartData(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.getRow(0).createCell(2).setCellValue("Actual");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.getRow(1).createCell(2).setCellValue(12d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(18d);
    sheet.getRow(2).createCell(2).setCellValue(16d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
    sheet.getRow(3).createCell(2).setCellValue(21d);
  }

  private static void removeFirstChartRelationship(Path workbookPath) throws IOException {
    try (var fileSystem = java.nio.file.FileSystems.newFileSystem(workbookPath)) {
      Path relationshipsPath = fileSystem.getPath("/xl/drawings/_rels/drawing1.xml.rels");
      String relationships = Files.readString(relationshipsPath);
      String updatedRelationships =
          relationships.replaceFirst("<Relationship[^>]+Type=\"[^\"]+/chart\"[^>]*/>", "");
      Files.writeString(relationshipsPath, updatedRelationships);
    }
  }

  private static <T extends ExcelChartSnapshot> T findChart(
      List<ExcelChartSnapshot> charts, String name, Class<T> type) {
    return type.cast(
        charts.stream().filter(snapshot -> name.equals(snapshot.name())).findFirst().orElseThrow());
  }
}
