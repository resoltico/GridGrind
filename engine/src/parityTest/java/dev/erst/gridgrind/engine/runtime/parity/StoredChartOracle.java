package dev.erst.gridgrind.engine.runtime.parity;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingCoordinateReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Optional;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

/** Independent stored-XML oracle; deliberately bypasses XDDF point and cache accessors. */
final class StoredChartOracle {
  private StoredChartOracle() {}

  static ChartReport.DataSource source(XmlObject source) {
    if (source == null) return new ChartReport.DataSource.Missing();
    Node root = source.getDomNode();
    for (Element choice : elements(root)) {
      switch (choice.getLocalName()) {
        case "numRef":
          return new ChartReport.DataSource.NumericReference(
              text(choice, "f"),
              child(choice, "numCache")
                  .<ChartReport.NumericCache>map(StoredChartOracle::numbers)
                  .orElseGet(ChartReport.NumericCache.Missing::new));
        case "strRef":
          return new ChartReport.DataSource.StringReference(
              text(choice, "f"),
              child(choice, "strCache")
                  .<ChartReport.StringCache>map(StoredChartOracle::strings)
                  .orElseGet(ChartReport.StringCache.Missing::new));
        case "numLit":
          return new ChartReport.DataSource.NumericLiteral(numbers(choice));
        case "strLit":
          return new ChartReport.DataSource.StringLiteral(strings(choice));
        case "multiLvlStrRef":
          return ChartReport.DataSource.Unsupported.unmodeled("MULTI_LEVEL_STRING_REFERENCE");
        default:
          throw new IllegalStateException("Unrecognized oracle chart source choice");
      }
    }
    return new ChartReport.DataSource.Missing();
  }

  static ChartReport.Title formula(XmlObject reference) {
    Node root = reference.getDomNode();
    return new ChartReport.Title.Formula(
        text(root, "f"),
        child(root, "strCache")
            .<ChartReport.StringCache>map(StoredChartOracle::strings)
            .orElseGet(ChartReport.StringCache.Missing::new));
  }

  private static ChartReport.StringCache.Stored strings(Element cache) {
    List<ChartReport.StringPoint> points = new ArrayList<>();
    for (Element point : elements(cache))
      if ("pt".equals(point.getLocalName())) {
        points.add(new ChartReport.StringPoint(attributeLong(point, "idx"), text(point, "v")));
      }
    return new ChartReport.StringCache.Stored(count(cache), points);
  }

  private static ChartReport.NumericCache.Stored numbers(Element cache) {
    List<ChartReport.NumericPoint> points = new ArrayList<>();
    for (Element point : elements(cache))
      if ("pt".equals(point.getLocalName())) {
        points.add(
            new ChartReport.NumericPoint(
                attributeLong(point, "idx"),
                text(point, "v"),
                point.hasAttribute("formatCode")
                    ? Optional.of(point.getAttribute("formatCode"))
                    : Optional.empty()));
      }
    return new ChartReport.NumericCache.Stored(count(cache), text(cache, "formatCode"), points);
  }

  private static Optional<Long> count(Element cache) {
    return child(cache, "ptCount").flatMap(value -> attributeLong(value, "val"));
  }

  private static Optional<Long> attributeLong(Element element, String name) {
    return element.hasAttribute(name)
        ? Optional.of(Long.parseLong(element.getAttribute(name)))
        : Optional.empty();
  }

  private static Optional<String> text(Node parent, String name) {
    return child(parent, name)
        .map(element -> {
          StringBuilder text = new StringBuilder();
          for (Node value = element.getFirstChild();
              value != null;
              value = value.getNextSibling()) {
            if (value.getNodeType() != Node.TEXT_NODE
                && value.getNodeType() != Node.CDATA_SECTION_NODE) {
              throw new IllegalStateException("Oracle scalar contains non-text content");
            }
            text.append(value.getNodeValue());
          }
          return text.toString();
        });
  }

  private static Optional<Element> child(Node parent, String name) {
    var matching =
        elements(parent).stream().filter(element -> name.equals(element.getLocalName())).toList();
    if (matching.size() > 1)
      throw new IllegalStateException("Repeated oracle singleton declaration");
    return matching.stream().findFirst();
  }

  private static List<Element> elements(Node parent) {
    List<Element> elements = new ArrayList<>();
    for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
      if (node instanceof Element element) elements.add(element);
    }
    return elements;
  }

  static List<ChartReport.Axis> axes(CTPlotArea area, List<CTUnsignedInt> references) {
    var indexed = new LinkedHashMap<Long, Element>();
    for (String family : List.of("catAx", "dateAx", "serAx", "valAx")) {
      for (Element axis : elements(area.getDomNode()))
        if (family.equals(axis.getLocalName())) {
          indexed.put(Long.valueOf(child(axis, "axId").orElseThrow().getAttribute("val")), axis);
        }
    }
    List<ChartReport.Axis> result = new ArrayList<>();
    for (CTUnsignedInt reference : references) {
      Element axis = indexed.get(reference.getVal());
      if (axis == null) continue;
      var kind =
          switch (axis.getLocalName()) {
            case "catAx" -> ExcelChartAxisKind.CATEGORY;
            case "dateAx" -> ExcelChartAxisKind.DATE;
            case "serAx" -> ExcelChartAxisKind.SERIES;
            case "valAx" -> ExcelChartAxisKind.VALUE;
            default -> throw new IllegalStateException("Unrecognized oracle native axis");
          };
      var position =
          switch (child(axis, "axPos").orElseThrow().getAttribute("val")) {
            case "b" -> ExcelChartAxisPosition.BOTTOM;
            case "l" -> ExcelChartAxisPosition.LEFT;
            case "r" -> ExcelChartAxisPosition.RIGHT;
            case "t" -> ExcelChartAxisPosition.TOP;
            default -> throw new IllegalStateException("Unrecognized oracle axis position");
          };
      boolean deleted =
          child(axis, "delete")
              .map(value ->
                  !value.hasAttribute("val")
                      || "true".equals(value.getAttribute("val"))
                      || "1".equals(value.getAttribute("val")))
              .orElse(false);
      result.add(new ChartReport.Axis(kind, position, crossing(axis), !deleted));
    }
    return List.copyOf(result);
  }

  private static ChartReport.AxisCrossing crossing(Element axis) {
    var mode = child(axis, "crosses");
    var numeric = child(axis, "crossesAt");
    if (mode.isPresent() && numeric.isPresent())
      throw new IllegalStateException("Mixed oracle axis crossing");
    if (numeric.isPresent())
      return new ChartReport.AxisCrossing.Value(numeric.orElseThrow().getAttribute("val"));
    if (mode.isEmpty()) return new ChartReport.AxisCrossing.Missing();
    return new ChartReport.AxisCrossing.Mode(
        switch (mode.orElseThrow().getAttribute("val")) {
          case "autoZero" -> ExcelChartAxisCrosses.AUTO_ZERO;
          case "min" -> ExcelChartAxisCrosses.MIN;
          case "max" -> ExcelChartAxisCrosses.MAX;
          default -> throw new IllegalStateException("Unrecognized oracle axis crossing");
        });
  }

  static DrawingAnchorReport anchor(XmlObject frame) {
    Node root = frame.getDomNode().getParentNode();
    return switch (root.getLocalName()) {
      case "twoCellAnchor" ->
          new DrawingAnchorReport.TwoCell(
              marker(child(root, "from").orElseThrow()),
              marker(child(root, "to").orElseThrow()),
              behavior(((Element) root).getAttribute("editAs")));
      case "oneCellAnchor" -> {
        Element ext = child(root, "ext").orElseThrow();
        yield new DrawingAnchorReport.OneCell(
            marker(child(root, "from").orElseThrow()),
            Long.parseLong(ext.getAttribute("cx")),
            Long.parseLong(ext.getAttribute("cy")),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      }
      case "absoluteAnchor" -> {
        Element pos = child(root, "pos").orElseThrow();
        Element ext = child(root, "ext").orElseThrow();
        yield new DrawingAnchorReport.Absolute(
            coordinate(pos.getAttribute("x")),
            coordinate(pos.getAttribute("y")),
            Long.parseLong(ext.getAttribute("cx")),
            Long.parseLong(ext.getAttribute("cy")),
            ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE);
      }
      default -> throw new IllegalStateException("Unrecognized oracle drawing anchor");
    };
  }

  private static DrawingMarkerReport marker(Element marker) {
    return new DrawingMarkerReport(
        Integer.parseInt(text(marker, "col").orElseThrow()),
        Integer.parseInt(text(marker, "row").orElseThrow()),
        coordinate(text(marker, "colOff").orElseThrow()),
        coordinate(text(marker, "rowOff").orElseThrow()));
  }

  private static DrawingCoordinateReport coordinate(String text) {
    return text.matches("-?[0-9]+")
        ? new DrawingCoordinateReport.Emu(Long.parseLong(text))
        : new DrawingCoordinateReport.Measured(text);
  }

  private static ExcelDrawingAnchorBehavior behavior(String token) {
    return switch (token) {
      case "", "twoCell" -> ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE;
      case "oneCell" -> ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE;
      case "absolute" -> ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE;
      default -> throw new IllegalStateException("Unrecognized oracle anchor behavior");
    };
  }
}
