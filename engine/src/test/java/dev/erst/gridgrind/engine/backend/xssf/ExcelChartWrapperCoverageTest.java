package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.engine.charts.ChartAdapter;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import java.io.IOException;
import java.lang.reflect.Proxy;
import java.util.List;
import java.util.Optional;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/** Focused coverage for no-runtime chart wrapper overloads and fallback helper seams. */
class ExcelChartWrapperCoverageTest {
  private static final String CHART_NAMESPACE =
      "http://schemas.openxmlformats.org/drawingml/2006/chart";
  private static final String TEST_NAMESPACE = "urn:gridgrind:test";
  private static final String TEST_CHART_QNAME = "gg:chart";
  private static final String TEST_NON_CHART_QNAME = "gg:not-a-chart";

  @Test
  void chartMutationAndSnapshotWrappersDelegateWithoutAnExplicitFormulaRuntime()
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ResolvedChart definition = lineChartDefinition("OpsChart");

      ChartAdapter.validateChart(sheet, definition);
      ChartReport.Title preparedTitle =
          ChartAdapter.prepareSeriesTitle(sheet, new ChartTitleSpec.Formula("B1"));
      ChartReport.Title.Formula resolvedSeriesTitle =
          assertInstanceOf(
              ChartReport.Title.Formula.class,
              ChartAdapter.snapshotSeriesTitle(
                  sheet,
                  formulaSeriesTitle("Charts!$B$1"),
                  FormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator())));

      assertInstanceOf(ChartReport.Title.Formula.class, preparedTitle);
      assertEquals("Plan", resolvedSeriesTitle.cachedText());

      ChartAdapter.createChart(sheet, definition);
      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFChart chart = drawing.getCharts().getFirst();
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      assertEquals(chart, ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
      assertEquals(1, ChartAdapter.snapshotPlots(chart, chart.getChartSeries()).size());
      assertEquals(1, ChartAdapter.snapshotPlots(chart, chart.getChartSeries(), null).size());

      assertEquals("OpsChart", ChartAdapter.snapshotChart(chart, graphicFrame).name());
      assertEquals("OpsChart", ChartAdapter.snapshotChartDrawingObject(chart, graphicFrame).name());
      assertInstanceOf(
          DrawingObjectReport.Chart.class, DrawingAdapter.snapshot(drawing, graphicFrame));
    }
  }

  @Test
  void controllerSheetSupportAndFacadeWrappersCoverDefaultRuntimeDelegation() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet controllerSheet = workbook.createSheet("Controller");
      ExcelChartTestSupport.seedChartData(controllerSheet);
      DrawingAdapter controller = new DrawingAdapter();
      controller.setChart(controllerSheet, lineChartDefinition("ControllerChart"));

      SheetDrawings support = new SheetDrawings(controllerSheet, controller);
      assertEquals(1, controller.charts(controllerSheet).size());
      assertEquals(1, support.charts().size());

      assertEquals(
          List.of("cached-only"),
          ChartAdapter.resolvedOrCachedReferenceValues(
              controllerSheet, " ", new ReferenceDataSource(" ", List.of("cached-only")), null));

      XSSFSheet facadeSheet = workbook.createSheet("Facade");
      ExcelChartTestSupport.seedChartData(facadeSheet);
      ResolvedChart facadeChart = lineChartDefinition("FacadeChart");
      ChartAdapter.validateChart(facadeSheet, facadeChart);
      ChartAdapter.createChart(facadeSheet, facadeChart);

      assertEquals(1, new DrawingAdapter().charts(facadeSheet).size());
    }
  }

  @Test
  void chartLookupReturnsTheLiveChartForWellFormedGraphicFrames() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("OpsChart"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFChart chart = drawing.getCharts().getFirst();
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();

      assertEquals(chart, ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
      assertNull(ChartAdapter.chartForGraphicFrame(drawing, null));
      assertEquals(1, ChartAdapter.snapshotPlots(chart, chart.getChartSeries()).size());
    }
  }

  @Test
  void chartLookupReturnsNullWhenGraphicDataIsMissing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("NoGraphic");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("NoGraphic"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      graphicFrame
          .getCTGraphicalObjectFrame()
          .getGraphic()
          .setGraphicData(
              org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData.Factory
                  .newInstance());

      assertNull(ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
    }
  }

  @Test
  void chartLookupReturnsNullWhenRelationIdIsMissingBlankOrUnknown() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("MissingRelation");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("MissingRelation"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      chartRelationNode(graphicFrame).getAttributes().removeNamedItem("r:id");

      assertNull(ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("BlankRelation");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("BlankRelation"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      ((Element) chartRelationNode(graphicFrame))
          .setAttributeNS(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", " ");

      assertNull(ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("UnknownRelation");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("UnknownRelation"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      ((Element) chartRelationNode(graphicFrame))
          .setAttributeNS(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
              "r:id",
              "rIdMissing");

      assertNull(ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
    }
  }

  @Test
  void chartLookupIgnoresMalformedGraphicDataChildren() throws Throwable {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("NonChartNode");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("NonChartNode"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      Node graphicData =
          graphicFrame.getCTGraphicalObjectFrame().getGraphic().getGraphicData().getDomNode();
      while (graphicData.hasChildNodes()) {
        graphicData.removeChild(graphicData.getFirstChild());
      }
      graphicData.appendChild(
          graphicData.getOwnerDocument().createElementNS(TEST_NAMESPACE, TEST_NON_CHART_QNAME));

      assertNull(ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
      assertEquals(Optional.empty(), ChartAdapter.chartRelationId(graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("LeadingNoiseNode");
      ExcelChartTestSupport.seedChartData(sheet);
      ChartAdapter.createChart(sheet, lineChartDefinition("LeadingNoiseNode"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFChart chart = drawing.getCharts().getFirst();
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      Node graphicData =
          graphicFrame.getCTGraphicalObjectFrame().getGraphic().getGraphicData().getDomNode();
      graphicData.insertBefore(
          graphicData.getOwnerDocument().createComment("noise"), graphicData.getFirstChild());

      assertEquals(chart, ChartAdapter.chartForGraphicFrame(drawing, graphicFrame));
      assertEquals(Node.COMMENT_NODE, graphicData.getFirstChild().getNodeType());
      assertTrue(ChartAdapter.chartRelationId(graphicFrame).orElseThrow().startsWith("rId"));
    }
  }

  @Test
  void chartNodeDetectionHelperRecognizesOnlyChartNodes() throws Exception {
    var document = DocumentBuilderFactory.newDefaultInstance().newDocumentBuilder().newDocument();

    assertTrue(ChartAdapter.isChartNode(document.createElementNS(CHART_NAMESPACE, "c:chart")));
    assertTrue(ChartAdapter.isChartNode(document.createElement("c:chart")));
    assertFalse(
        ChartAdapter.isChartNode(document.createElementNS(TEST_NAMESPACE, TEST_CHART_QNAME)));
    assertFalse(ChartAdapter.isChartNode(null));
  }

  @Test
  void graphicFrameRelationHelperHandlesMissingGraphicContainers() {
    assertEquals(Optional.empty(), ChartAdapter.chartRelationId((XSSFGraphicFrame) null));

    CTGraphicalObjectFrame missingGraphic = CTGraphicalObjectFrame.Factory.newInstance();
    assertEquals(Optional.empty(), ChartAdapter.chartRelationId(missingGraphic));

    CTGraphicalObjectFrame missingGraphicData = CTGraphicalObjectFrame.Factory.newInstance();
    missingGraphicData.addNewGraphic();
    assertEquals(Optional.empty(), ChartAdapter.chartRelationId(missingGraphicData));
  }

  @Test
  void nodeAndAttributeRelationHelpersTolerateNoiseAndMissingAttributes() throws Exception {
    var document = DocumentBuilderFactory.newDefaultInstance().newDocumentBuilder().newDocument();
    Element chartNode = document.createElementNS(CHART_NAMESPACE, "c:chart");

    assertEquals(Optional.empty(), ChartAdapter.chartRelationId(document.createComment("noise")));
    assertEquals(Optional.empty(), ChartAdapter.relationAttributeValue(null));
    assertEquals(Optional.empty(), ChartAdapter.chartRelationId(chartNode));

    chartNode.setAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", " ");
    assertEquals(Optional.empty(), ChartAdapter.chartRelationId(chartNode));

    chartNode.setAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", "rId7");
    assertEquals(Optional.of("rId7"), ChartAdapter.chartRelationId(chartNode));
    assertEquals(
        Optional.of("rId7"), ChartAdapter.relationAttributeValue(chartNode.getAttributes()));
    assertEquals(Optional.empty(), ChartAdapter.chartRelationId(attributeLessChartNode()));
  }

  @Test
  void replacingAFormulaTitledChartWithAnUntitledSeriesPersistsTheAbsence() throws Exception {
    var path = XlsxRoundTrip.newWorkbookPath("gridgrind-chart-series-title-");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet.xssfSheet());
      ResolvedChart original = lineChartDefinition("OpsChart");
      sheet.drawings().setChart(original);
      var first =
          assertInstanceOf(
              ChartReport.Line.class, sheet.drawings().charts().getFirst().plots().getFirst());
      assertInstanceOf(ChartReport.Title.Formula.class, first.series().getFirst().title());
      ResolvedChart untitled =
          ExcelChartTestSupport.lineChart(
              original.name(),
              original.anchor(),
              original.title(),
              original.legend(),
              original.displayBlanksAs(),
              original.plotOnlyVisibleCells(),
              false,
              List.of(
                  new ResolvedChart.Series(
                      new ChartTitleSpec.None(),
                      ExcelChartTestSupport.ref("A2:A4"),
                      ExcelChartTestSupport.ref("B2:B4"),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty())));
      sheet.drawings().setChart(untitled);
      var replaced =
          assertInstanceOf(
              ChartReport.Line.class, sheet.drawings().charts().getFirst().plots().getFirst());
      assertInstanceOf(ChartReport.Title.None.class, replaced.series().getFirst().title());
      workbook
          .persistence()
          .save(
              path,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    try (XssfWorkbook reopened =
        XssfWorkbooks.open(path, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      var plot =
          assertInstanceOf(
              ChartReport.Line.class,
              reopened.sheet("Charts").drawings().charts().getFirst().plots().getFirst());
      assertInstanceOf(ChartReport.Title.None.class, plot.series().getFirst().title());
      assertFalse(
          reopened
              .xssfWorkbook()
              .getSheet("Charts")
              .getDrawingPatriarch()
              .getCharts()
              .getFirst()
              .getCTChart()
              .getPlotArea()
              .getLineChartArray(0)
              .getSerArray(0)
              .isSetTx());
    }
  }

  @Test
  void formulaSeriesTitlesPreserveQuotedSheetReferencesAndCachedText() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet target = workbook.createSheet("Charts");
      XSSFSheet source = workbook.createSheet("Plan's Data");
      ExcelChartTestSupport.seedChartData(target);
      ExcelChartTestSupport.seedChartData(source);
      ChartAdapter.createChart(
          target, lineChartDefinition("Quoted", new ChartTitleSpec.Formula("'Plan''s Data'!$B$1")));
      var title =
          target
              .getDrawingPatriarch()
              .getCharts()
              .getFirst()
              .getCTChart()
              .getPlotArea()
              .getLineChartArray(0)
              .getSerArray(0)
              .getTx()
              .getStrRef();
      assertEquals("'Plan''s Data'!$B$1", title.getF());
      assertEquals("Plan", title.getStrCache().getPtArray(0).getV());
    }
  }

  private static ResolvedChart lineChartDefinition(String name) {
    return lineChartDefinition(name, new ChartTitleSpec.Formula("B1"));
  }

  private static ResolvedChart lineChartDefinition(String name, ResolvedChartTitle seriesTitle) {
    return ExcelChartTestSupport.lineChart(
        name,
        ExcelChartTestSupport.anchor(1, 5, 10, 18),
        new ResolvedChartTitle.Text("Ops"),
        new ChartLegendSpec.Hidden(),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        false,
        List.of(
            new ResolvedChart.Series(
                seriesTitle,
                ExcelChartTestSupport.ref("A2:A4"),
                ExcelChartTestSupport.ref("B2:B4"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())));
  }

  private static CTSerTx formulaSeriesTitle(String formula) {
    CTSerTx title = CTSerTx.Factory.newInstance();
    title.addNewStrRef().setF(formula);
    return title;
  }

  private static Node chartRelationNode(XSSFGraphicFrame graphicFrame) {
    NodeList children =
        graphicFrame
            .getCTGraphicalObjectFrame()
            .getGraphic()
            .getGraphicData()
            .getDomNode()
            .getChildNodes();
    for (int index = 0; index < children.getLength(); index++) {
      Node child = children.item(index);
      if (child.getNodeType() == Node.ELEMENT_NODE) {
        return child;
      }
    }
    throw new IllegalStateException("Expected chart graphic data to contain an element child");
  }

  private static Node attributeLessChartNode() {
    ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
    if (classLoader == null) {
      classLoader = ClassLoader.getPlatformClassLoader();
    }
    return (Node)
        Proxy.newProxyInstance(
            classLoader,
            new Class<?>[] {Node.class},
            (proxy, method, args) ->
                switch (method.getName()) {
                  case "getLocalName" -> "chart";
                  case "getNamespaceURI" -> CHART_NAMESPACE;
                  case "getNodeName" -> "c:chart";
                  case "getAttributes" -> null;
                  case "getNodeType" -> Node.ELEMENT_NODE;
                  case "hashCode" -> System.identityHashCode(proxy);
                  case "toString" -> "AttributeLessChartNode";
                  default -> defaultValue(method.getReturnType());
                });
  }

  private static Object defaultValue(Class<?> returnType) {
    if (!returnType.isPrimitive()) {
      return null;
    }
    if (returnType == boolean.class) {
      return false;
    }
    if (returnType == byte.class) {
      return (byte) 0;
    }
    if (returnType == short.class) {
      return (short) 0;
    }
    if (returnType == int.class) {
      return 0;
    }
    if (returnType == long.class) {
      return 0L;
    }
    if (returnType == float.class) {
      return 0F;
    }
    if (returnType == double.class) {
      return 0D;
    }
    if (returnType == char.class) {
      return '\0';
    }
    throw new IllegalArgumentException("Unsupported primitive type: " + returnType);
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
