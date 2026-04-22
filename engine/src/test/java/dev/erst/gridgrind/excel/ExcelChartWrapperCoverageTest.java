package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.reflect.Proxy;
import java.util.List;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/** Focused coverage for no-runtime chart wrapper overloads and fallback helper seams. */
class ExcelChartWrapperCoverageTest {
  private static final String CHART_NAMESPACE =
      "http://schemas.openxmlformats.org/drawingml/2006/chart";
  private static final String TEST_NAMESPACE = "urn:gridgrind:test";
  private static final String TEST_CHART_QNAME = "gg:chart";
  private static final String TEST_NON_CHART_QNAME = "gg:not-a-chart";
  private static final MethodHandle GRAPHIC_FRAME_RELATION_ID_HANDLE =
      graphicFrameRelationIdHandle();
  private static final MethodHandle NODE_RELATION_ID_HANDLE = nodeRelationIdHandle();
  private static final MethodHandle GRAPHIC_FRAME_CONSTRUCTOR_HANDLE =
      graphicFrameConstructorHandle();
  private static final MethodHandle IS_CHART_NODE_HANDLE = isChartNodeHandle();
  private static final MethodHandle RELATION_ATTRIBUTE_VALUE_HANDLE =
      relationAttributeValueHandle();

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
      assertEquals(chart, ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
      assertEquals(
          1, ExcelChartPlotSnapshotSupport.snapshotPlots(chart, chart.getChartSeries()).size());
      assertEquals(
          1,
          ExcelChartPlotSnapshotSupport.snapshotPlots(chart, chart.getChartSeries(), null).size());

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

  @Test
  void chartLookupReturnsTheLiveChartForWellFormedGraphicFrames() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("OpsChart"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFChart chart = drawing.getCharts().getFirst();
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();

      assertEquals(chart, ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, null));
      assertEquals(
          1, ExcelChartPlotSnapshotSupport.snapshotPlots(chart, chart.getChartSeries()).size());
    }
  }

  @Test
  void chartLookupReturnsNullWhenGraphicDataIsMissing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("NoGraphic");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("NoGraphic"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      graphicFrame
          .getCTGraphicalObjectFrame()
          .getGraphic()
          .setGraphicData(
              org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData.Factory
                  .newInstance());

      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
    }
  }

  @Test
  void chartLookupReturnsNullWhenRelationIdIsMissingBlankOrUnknown() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("MissingRelation");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("MissingRelation"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      chartRelationNode(graphicFrame).getAttributes().removeNamedItem("r:id");

      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("BlankRelation");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("BlankRelation"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      ((Element) chartRelationNode(graphicFrame))
          .setAttributeNS(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", " ");

      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("UnknownRelation");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("UnknownRelation"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      ((Element) chartRelationNode(graphicFrame))
          .setAttributeNS(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
              "r:id",
              "rIdMissing");

      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
    }
  }

  @Test
  void chartLookupIgnoresMalformedGraphicDataChildren() throws Throwable {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("NonChartNode");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("NonChartNode"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFGraphicFrame graphicFrame = drawing.getCharts().getFirst().getGraphicFrame();
      Node graphicData =
          graphicFrame.getCTGraphicalObjectFrame().getGraphic().getGraphicData().getDomNode();
      while (graphicData.hasChildNodes()) {
        graphicData.removeChild(graphicData.getFirstChild());
      }
      graphicData.appendChild(
          graphicData.getOwnerDocument().createElementNS(TEST_NAMESPACE, TEST_NON_CHART_QNAME));

      assertNull(ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
      assertNull(chartRelationIdViaHandle(graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("LeadingNoiseNode");
      ExcelChartTestSupport.seedChartData(sheet);
      ExcelChartMutationSupport.createChart(sheet, lineChartDefinition("LeadingNoiseNode"));

      XSSFDrawing drawing = sheet.getDrawingPatriarch();
      XSSFChart chart = drawing.getCharts().getFirst();
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      Node graphicData =
          graphicFrame.getCTGraphicalObjectFrame().getGraphic().getGraphicData().getDomNode();
      graphicData.insertBefore(
          graphicData.getOwnerDocument().createComment("noise"), graphicData.getFirstChild());

      assertEquals(chart, ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame));
      assertEquals(Node.COMMENT_NODE, graphicData.getFirstChild().getNodeType());
      assertTrue(chartRelationIdViaHandle(graphicFrame).startsWith("rId"));
    }
  }

  @Test
  void chartNodeDetectionHelperRecognizesOnlyChartNodes() throws Throwable {
    var document = DocumentBuilderFactory.newDefaultInstance().newDocumentBuilder().newDocument();

    assertTrue(isChartNodeViaHandle(document.createElementNS(CHART_NAMESPACE, "c:chart")));
    assertTrue(isChartNodeViaHandle(document.createElement("c:chart")));
    assertFalse(isChartNodeViaHandle(document.createElementNS(TEST_NAMESPACE, TEST_CHART_QNAME)));
    assertFalse(isChartNodeViaHandle(null));
  }

  @Test
  void graphicFrameRelationHelperHandlesMissingGraphicContainers() throws Throwable {
    assertNull(chartRelationIdViaHandle((XSSFGraphicFrame) null));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("GraphicMissing");
      ExcelChartTestSupport.seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFGraphicFrame graphicFrame =
          syntheticGraphicFrame(
              drawing,
              org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
                  .Factory.newInstance());

      assertNull(chartRelationIdViaHandle(graphicFrame));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("GraphicDataMissing");
      ExcelChartTestSupport.seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      var ctGraphicFrame =
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
              .Factory.newInstance();
      ctGraphicFrame.addNewGraphic();
      XSSFGraphicFrame graphicFrame = syntheticGraphicFrame(drawing, ctGraphicFrame);

      assertNull(chartRelationIdViaHandle(graphicFrame));
    }
  }

  @Test
  void nodeAndAttributeRelationHelpersTolerateNoiseAndMissingAttributes() throws Throwable {
    var document = DocumentBuilderFactory.newDefaultInstance().newDocumentBuilder().newDocument();
    Element chartNode = document.createElementNS(CHART_NAMESPACE, "c:chart");

    assertNull(nodeRelationIdViaHandle(document.createComment("noise")));
    assertNull(relationAttributeValueViaHandle(null));
    assertNull(nodeRelationIdViaHandle(chartNode));

    chartNode.setAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", " ");
    assertNull(nodeRelationIdViaHandle(chartNode));

    chartNode.setAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", "rId7");
    assertEquals("rId7", nodeRelationIdViaHandle(chartNode));
    assertEquals("rId7", relationAttributeValueViaHandle(chartNode.getAttributes()));
    assertNull(nodeRelationIdViaHandle(attributeLessChartNode()));
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

  private static MethodHandle graphicFrameRelationIdHandle() {
    try {
      return MethodHandles.privateLookupIn(ExcelChartSnapshotSupport.class, MethodHandles.lookup())
          .findStatic(
              ExcelChartSnapshotSupport.class,
              "chartRelationId",
              MethodType.methodType(String.class, XSSFGraphicFrame.class));
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("Failed to access chartRelationId helper", exception);
    }
  }

  private static MethodHandle nodeRelationIdHandle() {
    try {
      return MethodHandles.privateLookupIn(ExcelChartSnapshotSupport.class, MethodHandles.lookup())
          .findStatic(
              ExcelChartSnapshotSupport.class,
              "chartRelationId",
              MethodType.methodType(String.class, Node.class));
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("Failed to access chartRelationId node helper", exception);
    }
  }

  private static MethodHandle graphicFrameConstructorHandle() {
    try {
      return MethodHandles.privateLookupIn(XSSFGraphicFrame.class, MethodHandles.lookup())
          .findConstructor(
              XSSFGraphicFrame.class,
              MethodType.methodType(
                  void.class,
                  XSSFDrawing.class,
                  org.openxmlformats
                      .schemas
                      .drawingml
                      .x2006
                      .spreadsheetDrawing
                      .CTGraphicalObjectFrame
                      .class));
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("Failed to access XSSFGraphicFrame constructor", exception);
    }
  }

  private static MethodHandle isChartNodeHandle() {
    try {
      return MethodHandles.privateLookupIn(ExcelChartSnapshotSupport.class, MethodHandles.lookup())
          .findStatic(
              ExcelChartSnapshotSupport.class,
              "isChartNode",
              MethodType.methodType(boolean.class, Node.class));
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("Failed to access isChartNode helper", exception);
    }
  }

  private static MethodHandle relationAttributeValueHandle() {
    try {
      return MethodHandles.privateLookupIn(ExcelChartSnapshotSupport.class, MethodHandles.lookup())
          .findStatic(
              ExcelChartSnapshotSupport.class,
              "relationAttributeValue",
              MethodType.methodType(String.class, NamedNodeMap.class));
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("Failed to access relationAttributeValue helper", exception);
    }
  }

  private static String chartRelationIdViaHandle(XSSFGraphicFrame graphicFrame) throws Throwable {
    return (String) GRAPHIC_FRAME_RELATION_ID_HANDLE.invokeExact(graphicFrame);
  }

  private static String nodeRelationIdViaHandle(Node node) throws Throwable {
    return (String) NODE_RELATION_ID_HANDLE.invokeExact(node);
  }

  private static boolean isChartNodeViaHandle(Node node) throws Throwable {
    return (boolean) IS_CHART_NODE_HANDLE.invokeExact(node);
  }

  private static String relationAttributeValueViaHandle(NamedNodeMap attributes) throws Throwable {
    return (String) RELATION_ATTRIBUTE_VALUE_HANDLE.invokeExact(attributes);
  }

  private static XSSFGraphicFrame syntheticGraphicFrame(
      XSSFDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
          graphicFrame)
      throws Throwable {
    return (XSSFGraphicFrame) GRAPHIC_FRAME_CONSTRUCTOR_HANDLE.invokeExact(drawing, graphicFrame);
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
