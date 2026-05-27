package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.excel.DrawingObjectNotFoundException;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlObject;

/** Resolves named drawing objects to their backing POI shape and anchor metadata. */
final class ExcelDrawingObjectLocator {
  private ExcelDrawingObjectLocator() {}

  static ExcelDrawingController.LocatedShape requiredLocatedShape(
      XSSFSheet sheet, String objectName) {
    Optional<ExcelDrawingController.LocatedShape> located = optionalLocatedShape(sheet, objectName);
    if (located.isEmpty()) {
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    return located.orElseThrow();
  }

  static Optional<ExcelDrawingController.LocatedShape> optionalLocatedShape(
      XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    ExcelDrawingArgumentSupport.requireNonBlank(objectName, "objectName");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return Optional.empty();
    }
    XSSFShape matchedShape = null;
    XmlObject matchedShapeXml = null;
    for (XSSFShape shape : drawing.getShapes()) {
      if (!ExcelDrawingAnchorSupport.resolvedName(shape).equals(objectName)) {
        continue;
      }
      if (matchedShape != null) {
        throw ambiguousObjectName(sheet, objectName);
      }
      matchedShape = shape;
      matchedShapeXml = ExcelDrawingAnchorSupport.shapeXml(shape);
    }
    return matchedShape == null
        ? Optional.empty()
        : Optional.of(
            new ExcelDrawingController.LocatedShape(
                drawing,
                matchedShape,
                matchedShapeXml,
                ExcelDrawingAnchorSupport.parentAnchor(matchedShapeXml).orElse(null)));
  }

  static IllegalArgumentException ambiguousObjectName(XSSFSheet sheet, String objectName) {
    return new IllegalArgumentException(
        "Multiple drawing objects named '"
            + objectName
            + "' exist on sheet '"
            + sheet.getSheetName()
            + "'");
  }

  static IllegalArgumentException noBinaryPayloadException(XSSFSheet sheet, String objectName) {
    return new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' does not expose a binary payload");
  }
}
