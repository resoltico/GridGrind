package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

/** Shared anchor, shape-XML, and name helpers for drawing read and mutation workflows. */
final class ExcelDrawingAnchorSupport {
  private ExcelDrawingAnchorSupport() {}

  static org.apache.poi.xssf.usermodel.XSSFClientAnchor toPoiAnchor(
      XSSFDrawing drawing, ExcelDrawingAnchor.TwoCell anchor) {
    org.apache.poi.xssf.usermodel.XSSFClientAnchor poiAnchor =
        drawing.createAnchor(
            anchor.from().dx(),
            anchor.from().dy(),
            anchor.to().dx(),
            anchor.to().dy(),
            anchor.from().columnIndex(),
            anchor.from().rowIndex(),
            anchor.to().columnIndex(),
            anchor.to().rowIndex());
    poiAnchor.setAnchorType(toPoiBehavior(anchor.behavior()));
    return poiAnchor;
  }

  static ClientAnchor.AnchorType toPoiBehavior(ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE -> ClientAnchor.AnchorType.MOVE_AND_RESIZE;
      case MOVE_DONT_RESIZE -> ClientAnchor.AnchorType.MOVE_DONT_RESIZE;
      case DONT_MOVE_AND_RESIZE -> ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
    };
  }

  static ExcelDrawingAnchor snapshotAnchor(XmlObject shapeXml) {
    XmlObject parentAnchor = parentAnchor(shapeXml);
    return switch (parentAnchor) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
              twoCellAnchor ->
          new ExcelDrawingAnchor.TwoCell(
              marker(twoCellAnchor.getFrom()),
              marker(twoCellAnchor.getTo()),
              behavior(twoCellAnchor.getEditAs()));
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor
              oneCellAnchor ->
          new ExcelDrawingAnchor.OneCell(
              marker(oneCellAnchor.getFrom()),
              oneCellAnchor.getExt().getCx(),
              oneCellAnchor.getExt().getCy(),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor
              absoluteAnchor ->
          new ExcelDrawingAnchor.Absolute(
              ((Number) absoluteAnchor.getPos().getX()).longValue(),
              ((Number) absoluteAnchor.getPos().getY()).longValue(),
              absoluteAnchor.getExt().getCx(),
              absoluteAnchor.getExt().getCy(),
              ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE);
      case null -> throw new IllegalStateException("Drawing object is missing its parent anchor");
      default ->
          throw new IllegalStateException(
              "Unsupported parent anchor type: " + parentAnchor.getClass().getName());
    };
  }

  static ExcelDrawingAnchorBehavior behavior(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum editAs) {
    return switch (editAs == null
        ? org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.INT_TWO_CELL
        : editAs.intValue()) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.INT_ONE_CELL ->
          ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE;
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.INT_ABSOLUTE ->
          ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE;
      default -> ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE;
    };
  }

  static ExcelDrawingMarker marker(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker marker) {
    return new ExcelDrawingMarker(
        marker.getCol(),
        marker.getRow(),
        Math.toIntExact(((Number) marker.getColOff()).longValue()),
        Math.toIntExact(((Number) marker.getRowOff()).longValue()));
  }

  static XmlObject shapeXml(XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return picture.getCTPicture();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return objectData.getCTShape();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector connector) {
      return connector.getCTConnector();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
      return group.getCTGroupShape();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      return graphicFrame.getCTGraphicalObjectFrame();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return simpleShape.getCTShape();
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  @SuppressWarnings("PMD.UseTryWithResources")
  static XmlObject parentAnchor(XmlObject shapeXml) {
    org.apache.xmlbeans.XmlCursor cursor = shapeXml.newCursor();
    try {
      return cursor.toParent() ? cursor.getObject() : null;
    } finally {
      cursor.close();
    }
  }

  static String resolvedName(XSSFShape shape) {
    return nullIfBlank(shape.getShapeName()) != null ? shape.getShapeName() : defaultName(shape);
  }

  static int shapeType(String presetGeometryToken) {
    STShapeType.Enum preset = STShapeType.Enum.forString(presetGeometryToken);
    if (preset == null) {
      throw new IllegalArgumentException("Unsupported presetGeometryToken: " + presetGeometryToken);
    }
    return preset.intValue();
  }

  static void updateAnchorInPlace(
      XSSFSheet sheet,
      String objectName,
      XmlObject parentAnchor,
      ExcelDrawingAnchor.TwoCell anchor) {
    if (!(parentAnchor
        instanceof
        org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
            twoCellAnchor)) {
      throw new IllegalArgumentException(
          "Drawing object '"
              + objectName
              + "' on sheet '"
              + sheet.getSheetName()
              + "' is not backed by a two-cell anchor");
    }
    applyMarker(twoCellAnchor.getFrom(), anchor.from());
    applyMarker(twoCellAnchor.getTo(), anchor.to());
    twoCellAnchor.setEditAs(toPoiEditAs(anchor.behavior()));
  }

  static String defaultName(XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return "Picture-" + picture.getCTPicture().getNvPicPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return "Object-" + objectData.getCTShape().getNvSpPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector connector) {
      return "Connector-" + connector.getCTConnector().getNvCxnSpPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
      return "Group-" + group.getCTGroupShape().getNvGrpSpPr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      return "GraphicFrame-"
          + graphicFrame.getCTGraphicalObjectFrame().getNvGraphicFramePr().getCNvPr().getId();
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return "Shape-" + simpleShape.getCTShape().getNvSpPr().getCNvPr().getId();
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private static void applyMarker(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker target,
      ExcelDrawingMarker source) {
    target.setCol(source.columnIndex());
    target.setRow(source.rowIndex());
    target.setColOff(source.dx());
    target.setRowOff(source.dy());
  }

  static org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum toPoiEditAs(
      ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.TWO_CELL;
      case MOVE_DONT_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.ONE_CELL;
      case DONT_MOVE_AND_RESIZE ->
          org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.ABSOLUTE;
    };
  }

  static String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }
}
