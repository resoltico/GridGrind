package dev.erst.gridgrind.excel;

import java.util.Optional;
import javax.xml.namespace.QName;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/** XML relation and context lookup helpers for chart graphic frames. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelChartRelationSupport {
  private ExcelChartRelationSupport() {}

  public static @Nullable XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, @Nullable XSSFGraphicFrame graphicFrame) {
    return optionalChartForGraphicFrame(drawing, graphicFrame).orElse(null);
  }

  public static Optional<XSSFSheet> contextSheet(@Nullable XSSFGraphicFrame graphicFrame) {
    if (graphicFrame == null) {
      return Optional.empty();
    }
    return Optional.of(graphicFrame.getDrawing().getSheet());
  }

  public static @Nullable XSSFSheet contextSheet(
      @Nullable XSSFChart chart, @Nullable XSSFGraphicFrame graphicFrame) {
    XSSFGraphicFrame resolvedGraphicFrame =
        graphicFrame != null ? graphicFrame : chart == null ? null : chart.getGraphicFrame();
    return contextSheet(resolvedGraphicFrame).orElse(null);
  }

  public static Optional<String> chartRelationId(XSSFGraphicFrame graphicFrame) {
    return chartRelationNodes(graphicFrame).flatMap(ExcelChartRelationSupport::chartRelationId);
  }

  public static Optional<String> chartRelationId(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTGraphicalObjectFrame
          graphicFrame) {
    return chartRelationNodes(graphicFrame).flatMap(ExcelChartRelationSupport::chartRelationId);
  }

  public static Optional<NodeList> chartRelationNodes(@Nullable XSSFGraphicFrame graphicFrame) {
    return chartRelationNodes(
        graphicFrame == null ? null : graphicFrame.getCTGraphicalObjectFrame());
  }

  public static Optional<NodeList> chartRelationNodes(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.@Nullable CTGraphicalObjectFrame
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

  public static Optional<String> chartRelationId(Node node) {
    if (!isChartNode(node)) {
      return Optional.empty();
    }
    return relationAttributeValue(node.getAttributes());
  }

  public static Optional<String> relationAttributeValue(NamedNodeMap attributes) {
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

  public static boolean isChartNode(Node node) {
    return node != null
        && (("chart".equals(node.getLocalName())
                && "http://schemas.openxmlformats.org/drawingml/2006/chart"
                    .equals(node.getNamespaceURI()))
            || "c:chart".equals(node.getNodeName()));
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

  private static Optional<XSSFChart> optionalChartForGraphicFrame(
      XSSFDrawing drawing, @Nullable XSSFGraphicFrame graphicFrame) {
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
}
