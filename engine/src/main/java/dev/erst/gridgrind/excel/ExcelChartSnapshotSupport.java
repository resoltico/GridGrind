package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import javax.xml.namespace.QName;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/** Chart snapshot and chart-readback helpers. */
final class ExcelChartSnapshotSupport {
  private static final System.Logger LOGGER =
      System.getLogger(ExcelChartSnapshotSupport.class.getName());

  private ExcelChartSnapshotSupport() {}

  static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChartDrawingObject(chart, graphicFrame, null);
  }

  static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      ExcelFormulaRuntime formulaRuntime) {
    ExcelChartSnapshot snapshot = snapshotChart(chart, graphicFrame, formulaRuntime);
    boolean supported =
        snapshot.plots().stream().noneMatch(ExcelChartSnapshot.Unsupported.class::isInstance);
    return new ExcelDrawingObjectSnapshot.Chart(
        resolvedChartName(graphicFrame),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
        supported,
        chartPlotTypeTokens(chart),
        titleSummary(snapshot.title()));
  }

  static ExcelChartSnapshot snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChart(chart, graphicFrame, null);
  }

  static ExcelChartSnapshot snapshotChart(
      XSSFChart chart, XSSFGraphicFrame graphicFrame, ExcelFormulaRuntime formulaRuntime) {
    List<XDDFChartData> chartData = chart.getChartSeries();
    List<ExcelChartSnapshot.Plot> plots =
        ExcelChartPlotSnapshotSupport.snapshotPlots(chart, graphicFrame, chartData, formulaRuntime);
    return new ExcelChartSnapshot(
        resolvedChartName(graphicFrame),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
        snapshotTitle(chart, graphicFrame, formulaRuntime),
        snapshotLegend(chart),
        snapshotDisplayBlanks(chart),
        chart.isPlotOnlyVisibleCells(),
        plots);
  }

  static XSSFChart chartForGraphicFrame(XSSFDrawing drawing, XSSFGraphicFrame graphicFrame) {
    return optionalChartForGraphicFrame(drawing, graphicFrame).orElse(null);
  }

  static ExcelChartSnapshot.Title snapshotTitle(XSSFChart chart) {
    return snapshotTitle(chart, null);
  }

  static ExcelChartSnapshot.Title snapshotTitle(
      XSSFChart chart, ExcelFormulaRuntime formulaRuntime) {
    return snapshotTitle(chart, chart.getGraphicFrame(), formulaRuntime);
  }

  static ExcelChartSnapshot.Title snapshotTitle(
      XSSFChart chart, XSSFGraphicFrame graphicFrame, ExcelFormulaRuntime formulaRuntime) {
    if (!chart.getCTChart().isSetTitle()) {
      return new ExcelChartSnapshot.Title.None();
    }
    String formula = chart.getTitleFormula();
    if (formula != null) {
      return new ExcelChartSnapshot.Title.Formula(
          formula, cachedTitleText(chart, graphicFrame, formula, formulaRuntime));
    }
    String text = chart.getTitleText().getString();
    return text.isBlank()
        ? new ExcelChartSnapshot.Title.None()
        : new ExcelChartSnapshot.Title.Text(text);
  }

  static String cachedTitleText(XSSFChart chart, String formula) {
    return cachedTitleText(chart, formula, null);
  }

  static String cachedTitleText(
      XSSFChart chart, String formula, ExcelFormulaRuntime formulaRuntime) {
    return cachedTitleText(chart, chart.getGraphicFrame(), formula, formulaRuntime);
  }

  static String cachedTitleText(
      XSSFChart chart,
      XSSFGraphicFrame graphicFrame,
      String formula,
      ExcelFormulaRuntime formulaRuntime) {
    Optional<String> resolvedText =
        optionalResolvedTitleFormulaText(chart, graphicFrame, formula, formulaRuntime);
    if (resolvedText.isPresent()) {
      return resolvedText.orElseThrow();
    }
    if (!chart.getCTChart().isSetTitle()
        || !chart.getCTChart().getTitle().isSetTx()
        || !chart.getCTChart().getTitle().getTx().isSetStrRef()
        || !chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache()
        || chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().sizeOfPtArray() == 0) {
      return "";
    }
    String cachedText =
        chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV();
    return cachedText;
  }

  static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    return resolvedTitleFormulaText(chart, formula, null);
  }

  static String resolvedTitleFormulaText(
      XSSFChart chart, String formula, ExcelFormulaRuntime formulaRuntime) {
    return optionalResolvedTitleFormulaText(chart, chart.getGraphicFrame(), formula, formulaRuntime)
        .orElse("");
  }

  static Optional<String> optionalResolvedTitleFormulaText(
      XSSFChart chart,
      XSSFGraphicFrame graphicFrame,
      String formula,
      ExcelFormulaRuntime formulaRuntime) {
    try {
      XSSFGraphicFrame resolvedGraphicFrame =
          graphicFrame != null ? graphicFrame : chart == null ? null : chart.getGraphicFrame();
      return contextSheet(resolvedGraphicFrame)
          .map(
              sheet ->
                  ExcelChartSourceSupport.scalarText(
                      sheet,
                      ExcelChartSourceSupport.resolveSingleCellReference(
                          sheet, formula, "Chart title formula"),
                      formulaRuntime));
    } catch (IllegalArgumentException exception) {
      if (recoverableTitleFormulaResolutionFailure(exception)) {
        return Optional.empty();
      }
      LOGGER.log(
          System.Logger.Level.WARNING,
          "Failed to resolve chart title formula '" + formula + "'; using cached or empty title",
          exception);
      return Optional.empty();
    } catch (RuntimeException exception) {
      LOGGER.log(
          System.Logger.Level.WARNING,
          "Failed to resolve chart title formula '" + formula + "'; using cached or empty title",
          exception);
      return Optional.empty();
    }
  }

  static boolean barVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfBarChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getBarChartArray(0).getVaryColors());
  }

  static boolean lineVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfLineChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getLineChartArray(0).getVaryColors());
  }

  static boolean pieVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfPieChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getPieChartArray(0).getVaryColors());
  }

  static ExcelChartSnapshot.Title snapshotSeriesTitle(CTSerTx title) {
    return snapshotSeriesTitle(null, title, null);
  }

  static ExcelChartSnapshot.Title snapshotSeriesTitle(
      XSSFSheet contextSheet, CTSerTx title, ExcelFormulaRuntime formulaRuntime) {
    if (title == null) {
      return new ExcelChartSnapshot.Title.None();
    }
    if (title.isSetStrRef()) {
      if (contextSheet != null) {
        try {
          return new ExcelChartSnapshot.Title.Formula(
              title.getStrRef().getF(),
              ExcelChartSourceSupport.scalarText(
                  contextSheet,
                  ExcelChartSourceSupport.resolveSingleCellReference(
                      contextSheet, title.getStrRef().getF(), "Series title formula"),
                  formulaRuntime));
        } catch (RuntimeException ignored) {
          // Fall back to the embedded chart cache when the formula cannot be resolved live.
        }
      }
      String cachedText =
          title.getStrRef().isSetStrCache() && title.getStrRef().getStrCache().sizeOfPtArray() > 0
              ? title.getStrRef().getStrCache().getPtArray(0).getV()
              : "";
      return new ExcelChartSnapshot.Title.Formula(title.getStrRef().getF(), cachedText);
    }
    return title.isSetV()
        ? new ExcelChartSnapshot.Title.Text(title.getV())
        : new ExcelChartSnapshot.Title.None();
  }

  static String titleSummary(ExcelChartSnapshot.Title title) {
    return switch (title) {
      case ExcelChartSnapshot.Title.None _ -> "";
      case ExcelChartSnapshot.Title.Text text -> text.text();
      case ExcelChartSnapshot.Title.Formula formula ->
          formula.cachedText().isEmpty() ? formula.formula() : formula.cachedText();
    };
  }

  private static boolean recoverableTitleFormulaResolutionFailure(
      IllegalArgumentException exception) {
    return "Chart source formulas must not cache error values".equals(exception.getMessage());
  }

  static List<String> resolvedOrCachedReferenceValues(
      XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    return resolvedOrCachedReferenceValues(contextSheet, referenceFormula, source, null);
  }

  static List<String> resolvedOrCachedReferenceValues(
      XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSeriesSnapshotSupport.resolvedOrCachedReferenceValues(
        contextSheet, referenceFormula, source, formulaRuntime);
  }

  private static List<String> chartPlotTypeTokens(XSSFChart chart) {
    return chartPlotTypeTokens(chart.getChartSeries());
  }

  private static List<String> chartPlotTypeTokens(List<XDDFChartData> chartData) {
    List<String> tokens = new ArrayList<>();
    for (XDDFChartData value : chartData) {
      tokens.add(ExcelChartPoiBridge.plotTypeToken(value));
    }
    return List.copyOf(tokens);
  }

  static XSSFSheet contextSheet(XSSFChart chart, XSSFGraphicFrame graphicFrame) {
    XSSFGraphicFrame resolvedGraphicFrame =
        graphicFrame != null ? graphicFrame : chart == null ? null : chart.getGraphicFrame();
    return contextSheet(resolvedGraphicFrame).orElse(null);
  }

  static Optional<XSSFSheet> contextSheet(XSSFGraphicFrame graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    return Optional.of(graphicFrame.getDrawing().getSheet());
  }

  private static String resolvedChartName(XSSFGraphicFrame graphicFrame) {
    String name = ExcelChartSourceSupport.nullIfBlank(graphicFrame.getName());
    return name != null ? name : "Chart-" + graphicFrame.getId();
  }

  private static ExcelChartSnapshot.Legend snapshotLegend(XSSFChart chart) {
    if (!chart.getCTChart().isSetLegend()) {
      return new ExcelChartSnapshot.Legend.Hidden();
    }
    return new ExcelChartSnapshot.Legend.Visible(
        ExcelChartPoiBridge.fromPoiLegendPosition(
            new org.apache.poi.xddf.usermodel.chart.XDDFChartLegend(chart.getCTChart())
                .getPosition()));
  }

  private static ExcelChartDisplayBlanksAs snapshotDisplayBlanks(XSSFChart chart) {
    return chart.getCTChart().isSetDispBlanksAs()
        ? ExcelChartPoiBridge.fromPoiDisplayBlanks(chart.getCTChart().getDispBlanksAs().getVal())
        : ExcelChartDisplayBlanksAs.GAP;
  }

  private static boolean truthy(CTBoolean value) {
    return value != null && value.getVal();
  }

  static Optional<String> chartRelationId(XSSFGraphicFrame graphicFrame) {
    return chartRelationNodes(graphicFrame).flatMap(ExcelChartSnapshotSupport::chartRelationId);
  }

  static Optional<String> chartRelationId(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
          graphicFrame) {
    return chartRelationNodes(graphicFrame).flatMap(ExcelChartSnapshotSupport::chartRelationId);
  }

  private static Optional<String> chartRelationId(NodeList nodes) {
    for (int index = 0; index < nodes.getLength(); index++) {
      Optional<String> relationId = chartRelationId(nodes.item(index));
      if (relationId.isPresent()) {
        return relationId;
      }
    }
    return Optional.empty();
  }

  static Optional<NodeList> chartRelationNodes(XSSFGraphicFrame graphicFrame) {
    return chartRelationNodes(
        graphicFrame == null ? null : graphicFrame.getCTGraphicalObjectFrame());
  }

  static Optional<NodeList> chartRelationNodes(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
          graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    var graphic = graphicFrame.getGraphic();
    if (graphic == null || graphic.getGraphicData() == null) {
      return Optional.empty();
    }
    return Optional.of(graphic.getGraphicData().getDomNode().getChildNodes());
  }

  static Optional<String> chartRelationId(Node node) {
    if (!isChartNode(node)) {
      return Optional.empty();
    }
    return relationAttributeValue(node.getAttributes());
  }

  static Optional<String> relationAttributeValue(NamedNodeMap attributes) {
    if (attributes == null) {
      return Optional.empty();
    }
    Node relationAttribute =
        attributes.getNamedItemNS(
            QName.valueOf("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                .getNamespaceURI(),
            "id");
    if (relationAttribute == null) {
      relationAttribute = attributes.getNamedItem("r:id");
    }
    if (relationAttribute == null || relationAttribute.getNodeValue().isBlank()) {
      return Optional.empty();
    }
    return Optional.of(relationAttribute.getNodeValue());
  }

  private static Optional<XSSFChart> optionalChartForGraphicFrame(
      XSSFDrawing drawing, XSSFGraphicFrame graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    Optional<String> relationId = chartRelationId(graphicFrame);
    if (relationId.isEmpty()) {
      return Optional.empty();
    }
    POIXMLDocumentPart relation = drawing.getRelationById(relationId.orElseThrow());
    return relation instanceof XSSFChart chart ? Optional.of(chart) : Optional.empty();
  }

  static boolean isChartNode(Node node) {
    return node != null
        && (("chart".equals(node.getLocalName())
                && "http://schemas.openxmlformats.org/drawingml/2006/chart"
                    .equals(node.getNamespaceURI()))
            || "c:chart".equals(node.getNodeName()));
  }
}
