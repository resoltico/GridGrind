package dev.erst.gridgrind.excel;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

/** Snapshot helpers for drawing-object inventories and shape summaries. */
final class ExcelDrawingSnapshotSupport {
  private ExcelDrawingSnapshotSupport() {}

  static ExcelDrawingObjectSnapshot snapshot(XSSFDrawing drawing, XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return snapshotPicture(picture);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return ExcelDrawingBinarySupport.snapshotEmbeddedObject(objectData);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector connector) {
      return snapshotConnector(connector);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
      return snapshotGroup(drawing, group);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      XSSFChart chart = ExcelDrawingChartSupport.chartForGraphicFrame(drawing, graphicFrame);
      return chart == null
          ? snapshotGraphicFrame(graphicFrame)
          : ExcelDrawingChartSupport.snapshotChartDrawingObject(chart, graphicFrame);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape simpleShape) {
      return snapshotShape(simpleShape);
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  static ExcelDrawingObjectSnapshot.Shape snapshotShape(
      org.apache.poi.xssf.usermodel.XSSFSimpleShape shape) {
    STShapeType.Enum preset =
        !shape.getCTShape().getSpPr().isSetPrstGeom()
            ? null
            : shape.getCTShape().getSpPr().getPrstGeom().getPrst();
    return new ExcelDrawingObjectSnapshot.Shape(
        ExcelDrawingAnchorSupport.resolvedName(shape),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(shape)),
        ExcelDrawingShapeKind.SIMPLE_SHAPE,
        preset == null ? null : preset.toString(),
        ExcelDrawingBinarySupport.nullIfBlank(shape.getText()),
        0);
  }

  private static ExcelDrawingObjectSnapshot.Picture snapshotPicture(XSSFPicture picture) {
    org.apache.poi.xssf.usermodel.XSSFPictureData pictureData = picture.getPictureData();
    byte[] bytes = pictureData.getData();
    RasterDimensions dimensions = rasterDimensions(bytes);
    return new ExcelDrawingObjectSnapshot.Picture(
        ExcelDrawingAnchorSupport.resolvedName(picture),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(picture)),
        ExcelPictureFormat.fromPoiPictureType(pictureData.getPictureType()),
        pictureData.getPackagePart().getContentType(),
        bytes.length,
        ExcelDrawingBinarySupport.sha256(bytes),
        dimensions.widthPixels(),
        dimensions.heightPixels(),
        ExcelDrawingBinarySupport.nullIfBlank(
            picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()));
  }

  private static ExcelDrawingObjectSnapshot.Shape snapshotConnector(
      org.apache.poi.xssf.usermodel.XSSFConnector connector) {
    return new ExcelDrawingObjectSnapshot.Shape(
        ExcelDrawingAnchorSupport.resolvedName(connector),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(connector)),
        ExcelDrawingShapeKind.CONNECTOR,
        null,
        null,
        0);
  }

  private static ExcelDrawingObjectSnapshot.Shape snapshotGroup(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFShapeGroup group) {
    return new ExcelDrawingObjectSnapshot.Shape(
        ExcelDrawingAnchorSupport.resolvedName(group),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(group)),
        ExcelDrawingShapeKind.GROUP,
        null,
        null,
        drawing.getShapes(group).size());
  }

  private static ExcelDrawingObjectSnapshot.Shape snapshotGraphicFrame(
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return new ExcelDrawingObjectSnapshot.Shape(
        ExcelDrawingAnchorSupport.resolvedName(graphicFrame),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
        ExcelDrawingShapeKind.GRAPHIC_FRAME,
        null,
        null,
        0);
  }

  static RasterDimensions rasterDimensions(byte[] bytes) {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      BufferedImage image = javax.imageio.ImageIO.read(input);
      return image == null
          ? RasterDimensions.none()
          : new RasterDimensions(image.getWidth(), image.getHeight());
    } catch (IOException exception) {
      return RasterDimensions.none();
    }
  }

  record RasterDimensions(Integer widthPixels, Integer heightPixels) {
    static RasterDimensions none() {
      return new RasterDimensions(null, null);
    }
  }
}
