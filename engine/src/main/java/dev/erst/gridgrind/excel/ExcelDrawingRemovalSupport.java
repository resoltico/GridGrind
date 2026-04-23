package dev.erst.gridgrind.excel;

import java.util.function.BiPredicate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlObject;

/** Removal helpers for drawings, charts, pictures, and embedded objects. */
final class ExcelDrawingRemovalSupport {
  private ExcelDrawingRemovalSupport() {}

  static void deleteLocatedShape(
      XSSFSheet sheet,
      ExcelDrawingController.LocatedShape located,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      deletePicture(sheet, located, picture);
      return;
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
      XSSFChart chart =
          ExcelDrawingChartSupport.chartForGraphicFrame(located.drawing(), graphicFrame);
      if (chart != null) {
        deleteChart(located, chart, poiRelationRemover);
        return;
      }
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      deleteEmbeddedObject(sheet, located, objectData);
      return;
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFShapeGroup) {
      removeParentAnchor(located.drawing(), located.parentAnchor());
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + ExcelDrawingAnchorSupport.resolvedName(located.shape())
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  private static void deleteChart(
      ExcelDrawingController.LocatedShape located,
      XSSFChart chart,
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    if (!poiRelationRemover.test(located.drawing(), chart)) {
      throw new IllegalStateException(
          "Failed to remove chart relation for '"
              + ExcelDrawingAnchorSupport.resolvedName(
                  (org.apache.poi.xssf.usermodel.XSSFGraphicFrame) located.shape())
              + "'");
    }
    removeParentAnchor(located.drawing(), located.parentAnchor());
  }

  private static void deletePicture(
      XSSFSheet sheet, ExcelDrawingController.LocatedShape located, XSSFPicture picture) {
    ExcelDrawingPictureSupport.deletePicture(sheet, located, picture);
  }

  private static void deleteEmbeddedObject(
      XSSFSheet sheet,
      ExcelDrawingController.LocatedShape located,
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    org.apache.poi.openxml4j.opc.PackagePart objectPart =
        ExcelDrawingBinarySupport.relatedInternalPart(
            sheetPart(objectData), objectData.getOleObject().getId());
    org.apache.poi.openxml4j.opc.PackagePart previewPart =
        ExcelDrawingBinarySupport.previewImagePart(objectData);
    org.apache.poi.openxml4j.opc.PackagePartName olePartName =
        objectPart == null ? null : objectPart.getPartName();
    org.apache.poi.openxml4j.opc.PackagePartName previewPartName =
        previewPart == null ? null : previewPart.getPartName();
    String drawingPreviewRelationId =
        ExcelDrawingBinarySupport.previewDrawingRelationId(objectData);
    String sheetPreviewRelationId =
        ExcelDrawingBinarySupport.previewSheetImageRelationId(objectData);
    String oleRelationId = objectData.getOleObject().getId();

    if (drawingPreviewRelationId != null) {
      located.drawing().getPackagePart().removeRelationship(drawingPreviewRelationId);
    }
    if (sheetPreviewRelationId != null) {
      sheet.getPackagePart().removeRelationship(sheetPreviewRelationId);
    }
    if (oleRelationId != null) {
      sheet.getPackagePart().removeRelationship(oleRelationId);
    }
    ExcelDrawingBinarySupport.removeOleObject(sheet, objectData.getOleObject());
    removeParentAnchor(located.drawing(), located.parentAnchor());
    cleanupPackagePartIfUnused(sheet.getWorkbook().getPackage(), olePartName);
    ExcelDrawingBinarySupport.cleanupWorkbookImagePartIfUnused(
        sheet.getWorkbook(), previewPartName);
  }

  static void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
    ExcelPackageRelationshipSupport.cleanupPackagePartIfUnused(pkg, partName);
  }

  static void removeParentAnchor(XSSFDrawing drawing, XmlObject parentAnchor) {
    switch (parentAnchor) {
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor
              twoCellAnchor ->
          removeTwoCellAnchor(drawing.getCTDrawing(), twoCellAnchor);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor
              oneCellAnchor ->
          removeOneCellAnchor(drawing.getCTDrawing(), oneCellAnchor);
      case org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor
              absoluteAnchor ->
          removeAbsoluteAnchor(drawing.getCTDrawing(), absoluteAnchor);
      case null -> throw new IllegalStateException("Drawing object is missing its parent anchor");
      default ->
          throw new IllegalStateException(
              "Unsupported parent anchor type: " + parentAnchor.getClass().getName());
    }
  }

  static void removeTwoCellAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor target) {
    for (int index = 0; index < drawing.sizeOfTwoCellAnchorArray(); index++) {
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor candidate =
          drawing.getTwoCellAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeTwoCellAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate two-cell drawing anchor for deletion");
  }

  static void removeOneCellAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor target) {
    for (int index = 0; index < drawing.sizeOfOneCellAnchorArray(); index++) {
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor candidate =
          drawing.getOneCellAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeOneCellAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate one-cell drawing anchor for deletion");
  }

  static void removeAbsoluteAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor target) {
    for (int index = 0; index < drawing.sizeOfAbsoluteAnchorArray(); index++) {
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor candidate =
          drawing.getAbsoluteAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeAbsoluteAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate absolute drawing anchor for deletion");
  }

  private static org.apache.poi.openxml4j.opc.PackagePart sheetPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ((XSSFSheet) objectData.getDrawing().getParent()).getPackagePart();
  }
}
