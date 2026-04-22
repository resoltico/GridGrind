package dev.erst.gridgrind.excel;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.BiPredicate;
import javax.imageio.ImageIO;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;

/** Package-aware drawing and media controller for read and mutation workflows. */
final class ExcelDrawingController {
  private final BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover;
  private final ExcelSignatureLineController signatureLineController;

  ExcelDrawingController() {
    this(PoiRelationRemoval.defaultRemover());
  }

  ExcelDrawingController(BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover) {
    this(poiRelationRemover, new ExcelSignatureLineController());
  }

  ExcelDrawingController(
      BiPredicate<POIXMLDocumentPart, POIXMLDocumentPart> poiRelationRemover,
      ExcelSignatureLineController signatureLineController) {
    this.poiRelationRemover =
        Objects.requireNonNull(poiRelationRemover, "poiRelationRemover must not be null");
    this.signatureLineController =
        Objects.requireNonNull(signatureLineController, "signatureLineController must not be null");
  }

  List<ExcelDrawingObjectSnapshot> drawingObjects(XSSFSheet sheet) {
    return drawingObjects(sheet, null);
  }

  List<ExcelDrawingObjectSnapshot> drawingObjects(
      XSSFSheet sheet, ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    List<ExcelDrawingObjectSnapshot> snapshots = new ArrayList<>();
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing != null) {
      for (XSSFShape shape : drawing.getShapes()) {
        snapshots.add(snapshot(drawing, shape, formulaRuntime));
      }
    }
    for (ExcelSignatureLineSnapshot signatureLine : signatureLineController.signatureLines(sheet)) {
      snapshots.add(toDrawingObjectSnapshot(signatureLine));
    }
    return List.copyOf(snapshots);
  }

  List<ExcelChartSnapshot> charts(XSSFSheet sheet) {
    return charts(sheet, null);
  }

  List<ExcelChartSnapshot> charts(XSSFSheet sheet, ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return List.of();
    }
    drawing.getShapes();

    List<ExcelChartSnapshot> snapshots = new ArrayList<>();
    for (XSSFChart chart : drawing.getCharts()) {
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      if (graphicFrame == null) {
        continue;
      }
      snapshots.add(ExcelDrawingChartSupport.snapshotChart(chart, graphicFrame, formulaRuntime));
    }
    return List.copyOf(snapshots);
  }

  ExcelDrawingObjectPayload drawingObjectPayload(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");

    LocatedShape located = optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located != null && signatureLine) {
      throw ambiguousObjectName(sheet, objectName);
    }
    if (located == null) {
      if (signatureLine) {
        throw noBinaryPayloadException(sheet, objectName);
      }
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      return picturePayload(objectName, picture);
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      return embeddedObjectPayload(objectName, objectData);
    }
    throw noBinaryPayloadException(sheet, objectName);
  }

  void setPicture(XSSFSheet sheet, ExcelPictureDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    int pictureIndex =
        sheet
            .getWorkbook()
            .addPicture(definition.imageData().bytes(), definition.format().poiPictureType());
    XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);
    picture.getCTPicture().getNvPicPr().getCNvPr().setName(definition.name());
    if (definition.description() != null) {
      picture.getCTPicture().getNvPicPr().getCNvPr().setDescr(definition.description());
    }
  }

  void setChart(XSSFSheet sheet, ExcelChartDefinition definition) {
    setChart(sheet, definition, null);
  }

  void setChart(
      XSSFSheet sheet, ExcelChartDefinition definition, ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    ExcelDrawingChartSupport.validateChart(sheet, definition, formulaRuntime);
    deleteNamedObjectIfPresent(sheet, definition.name());
    try {
      ExcelDrawingChartSupport.createChart(sheet, definition, formulaRuntime);
    } catch (RuntimeException exception) {
      deleteNamedObjectIfPresent(sheet, definition.name());
      throw exception;
    }
  }

  void setSignatureLine(XSSFSheet sheet, ExcelSignatureLineDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    signatureLineController.setSignatureLine(sheet, definition);
  }

  void setShape(XSSFSheet sheet, ExcelShapeDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    Integer preparedShapeType =
        definition.kind() == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE
            ? shapeType(definition.presetGeometryToken())
            : null;
    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    if (definition.kind() == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE) {
      org.apache.poi.xssf.usermodel.XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
      shape.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
      shape.setShapeType(Objects.requireNonNull(preparedShapeType));
      if (definition.text() != null) {
        shape.setText(definition.text());
      }
      return;
    }
    org.apache.poi.xssf.usermodel.XSSFConnector connector = drawing.createConnector(anchor);
    connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(definition.name());
  }

  void setEmbeddedObject(XSSFSheet sheet, ExcelEmbeddedObjectDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedObjectIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        toPoiAnchor(drawing, definition.anchor());
    org.apache.poi.xssf.usermodel.XSSFObjectData objectData =
        ExcelIoSupport.unchecked(
            "Failed to create embedded object '" + definition.name() + "'",
            () -> {
              int storageId =
                  sheet
                      .getWorkbook()
                      .addOlePackage(
                          definition.payload().bytes(),
                          definition.label(),
                          definition.fileName(),
                          definition.command());
              int pictureIndex =
                  sheet
                      .getWorkbook()
                      .addPicture(
                          definition.previewImage().bytes(),
                          definition.previewFormat().poiPictureType());
              return drawing.createObjectData(anchor, storageId, pictureIndex);
            });
    objectData.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
  }

  void setDrawingObjectAnchor(
      XSSFSheet sheet, String objectName, ExcelDrawingAnchor.TwoCell anchor) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");

    LocatedShape located = optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located != null && signatureLine) {
      throw ambiguousObjectName(sheet, objectName);
    }
    if (signatureLine) {
      signatureLineController.updateAnchorIfPresent(sheet, objectName, anchor);
      return;
    }
    LocatedShape requiredLocated = requiredLocatedShape(sheet, objectName);
    XSSFShape shape = requiredLocated.shape();
    if (shape instanceof XSSFPicture
        || (shape instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame
            && ExcelDrawingChartSupport.chartForGraphicFrame(
                    requiredLocated.drawing(), graphicFrame)
                != null)
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFConnector
        || shape instanceof org.apache.poi.xssf.usermodel.XSSFSimpleShape) {
      updateAnchorInPlace(sheet, objectName, requiredLocated.parentAnchor(), anchor);
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  void deleteDrawingObject(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    boolean signatureLine = signatureLineController.hasNamedSignatureLine(sheet, objectName);
    if (located != null && signatureLine) {
      throw ambiguousObjectName(sheet, objectName);
    }
    if (located != null) {
      deleteLocatedShape(sheet, located);
      return;
    }
    if (signatureLineController.deleteIfPresent(sheet, objectName)) {
      return;
    }
    throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
  }

  void cleanupEmptyDrawingPatriarch(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    // POI exposes no supported API for unregistering a live XSSFDrawing relation object from its
    // parent sheet. Hard-deleting the package part leaves the in-memory relation graph stale and
    // breaks later save/commit flows on reopened workbooks. Prefer preserving an inert empty
    // drawing part over emitting a corrupt package.
  }

  private ExcelDrawingObjectSnapshot snapshot(
      XSSFDrawing drawing, XSSFShape shape, ExcelFormulaRuntime formulaRuntime) {
    return ExcelDrawingSnapshotSupport.snapshot(drawing, shape, formulaRuntime);
  }

  private ExcelDrawingObjectSnapshot.SignatureLine toDrawingObjectSnapshot(
      ExcelSignatureLineSnapshot signatureLine) {
    return new ExcelDrawingObjectSnapshot.SignatureLine(
        signatureLine.name(),
        signatureLine.anchor(),
        signatureLine.setupId(),
        signatureLine.allowComments(),
        signatureLine.signingInstructions(),
        signatureLine.suggestedSigner(),
        signatureLine.suggestedSigner2(),
        signatureLine.suggestedSignerEmail(),
        signatureLine.previewFormat(),
        signatureLine.previewContentType(),
        signatureLine.previewByteSize(),
        signatureLine.previewSha256(),
        signatureLine.previewWidthPixels(),
        signatureLine.previewHeightPixels());
  }

  ExcelDrawingObjectSnapshot.Shape snapshotShape(
      org.apache.poi.xssf.usermodel.XSSFSimpleShape shape) {
    return ExcelDrawingSnapshotSupport.snapshotShape(shape);
  }

  ExcelDrawingObjectSnapshot.EmbeddedObject snapshotEmbeddedObject(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ExcelDrawingBinarySupport.snapshotEmbeddedObject(objectData);
  }

  private ExcelDrawingObjectPayload.Picture picturePayload(String objectName, XSSFPicture picture) {
    return ExcelDrawingBinarySupport.picturePayload(objectName, picture);
  }

  private ExcelDrawingObjectPayload.EmbeddedObject embeddedObjectPayload(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ExcelDrawingBinarySupport.embeddedObjectPayload(objectName, objectData);
  }

  private void deleteNamedShapeIfPresent(XSSFSheet sheet, String objectName) {
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    if (located != null) {
      deleteLocatedShape(sheet, located);
    }
  }

  private void deleteNamedObjectIfPresent(XSSFSheet sheet, String objectName) {
    deleteNamedShapeIfPresent(sheet, objectName);
    signatureLineController.deleteIfPresent(sheet, objectName);
  }

  private void deleteLocatedShape(XSSFSheet sheet, LocatedShape located) {
    ExcelDrawingRemovalSupport.deleteLocatedShape(sheet, located, poiRelationRemover);
    cleanupEmptyDrawingPatriarch(sheet);
  }

  org.apache.poi.openxml4j.opc.PackagePart oleObjectPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId());
  }

  org.apache.poi.openxml4j.opc.PackagePart previewImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ExcelDrawingBinarySupport.previewImagePart(objectData);
  }

  void removeOleObject(
      XSSFSheet sheet, org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject target) {
    ExcelDrawingBinarySupport.removeOleObject(sheet, target);
  }

  void cleanupWorkbookImagePartIfUnused(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    ExcelDrawingBinarySupport.cleanupWorkbookImagePartIfUnused(workbook, imagePartName);
  }

  boolean imagePartUsed(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    return ExcelDrawingBinarySupport.imagePartUsed(workbook, imagePartName);
  }

  void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
    ExcelDrawingRemovalSupport.cleanupPackagePartIfUnused(pkg, partName);
  }

  org.apache.poi.openxml4j.opc.PackagePart relatedInternalPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart, String relationshipId) {
    return ExcelDrawingBinarySupport.relatedInternalPart(sourcePart, relationshipId);
  }

  void removeRelationshipsToPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart,
      org.apache.poi.openxml4j.opc.PackagePartName targetPartName) {
    ExcelDrawingBinarySupport.removeRelationshipsToPart(sourcePart, targetPartName);
  }

  private LocatedShape requiredLocatedShape(XSSFSheet sheet, String objectName) {
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    if (located == null) {
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    return located;
  }

  private IllegalArgumentException ambiguousObjectName(XSSFSheet sheet, String objectName) {
    return new IllegalArgumentException(
        "Multiple drawing objects named '"
            + objectName
            + "' exist on sheet '"
            + sheet.getSheetName()
            + "'");
  }

  private IllegalArgumentException noBinaryPayloadException(XSSFSheet sheet, String objectName) {
    return new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' does not expose a binary payload");
  }

  private LocatedShape optionalLocatedShape(XSSFSheet sheet, String objectName) {
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return null;
    }
    XSSFShape matchedShape = null;
    XmlObject matchedShapeXml = null;
    for (XSSFShape shape : drawing.getShapes()) {
      if (!resolvedName(shape).equals(objectName)) {
        continue;
      }
      if (matchedShape != null) {
        throw new IllegalArgumentException(
            "Multiple drawing objects named '"
                + objectName
                + "' exist on sheet '"
                + sheet.getSheetName()
                + "'");
      }
      matchedShape = shape;
      matchedShapeXml = shapeXml(shape);
    }
    return matchedShape == null
        ? null
        : new LocatedShape(drawing, matchedShape, matchedShapeXml, parentAnchor(matchedShapeXml));
  }

  static RasterDimensions rasterDimensions(byte[] bytes) {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      BufferedImage image = ImageIO.read(input);
      return image == null
          ? RasterDimensions.none()
          : new RasterDimensions(image.getWidth(), image.getHeight());
    } catch (IOException exception) {
      return RasterDimensions.none();
    }
  }

  /** Supported scalar kinds extracted from chart source cells. */
  enum CellScalarKind {
    STRING,
    NUMERIC
  }

  record CellScalar(CellScalarKind kind, String text, double number) {
    CellScalar {
      Objects.requireNonNull(kind, "kind must not be null");
      if (kind == CellScalarKind.STRING) {
        Objects.requireNonNull(text, "text must not be null");
      }
    }
  }

  private org.apache.poi.xssf.usermodel.XSSFClientAnchor toPoiAnchor(
      XSSFDrawing drawing, ExcelDrawingAnchor.TwoCell anchor) {
    return ExcelDrawingAnchorSupport.toPoiAnchor(drawing, anchor);
  }

  ClientAnchor.AnchorType toPoiBehavior(ExcelDrawingAnchorBehavior behavior) {
    return ExcelDrawingAnchorSupport.toPoiBehavior(behavior);
  }

  ExcelDrawingAnchor snapshotAnchor(XmlObject shapeXml) {
    return ExcelDrawingAnchorSupport.snapshotAnchor(shapeXml);
  }

  ExcelDrawingAnchorBehavior behavior(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum editAs) {
    return ExcelDrawingAnchorSupport.behavior(editAs);
  }

  String previewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject) {
    return ExcelDrawingBinarySupport.previewSheetRelationId(oleObject);
  }

  String previewDrawingRelationId(org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ExcelDrawingBinarySupport.previewDrawingRelationId(objectData);
  }

  XmlObject shapeXml(XSSFShape shape) {
    return ExcelDrawingAnchorSupport.shapeXml(shape);
  }

  XmlObject parentAnchor(XmlObject shapeXml) {
    return ExcelDrawingAnchorSupport.parentAnchor(shapeXml);
  }

  void removeParentAnchor(XSSFDrawing drawing, XmlObject parentAnchor) {
    ExcelDrawingRemovalSupport.removeParentAnchor(drawing, parentAnchor);
  }

  String resolvedName(XSSFShape shape) {
    return ExcelDrawingAnchorSupport.resolvedName(shape);
  }

  String defaultName(XSSFShape shape) {
    return ExcelDrawingAnchorSupport.defaultName(shape);
  }

  int shapeType(String presetGeometryToken) {
    return ExcelDrawingAnchorSupport.shapeType(presetGeometryToken);
  }

  ExcelBinaryData binary(byte[] bytes, String label) {
    return ExcelDrawingBinarySupport.binary(bytes, label);
  }

  byte[] partBytes(org.apache.poi.openxml4j.opc.PackagePart part) {
    return ExcelDrawingBinarySupport.partBytes(part);
  }

  String sha256(byte[] bytes) {
    return ExcelDrawingBinarySupport.sha256(bytes);
  }

  String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  String nullIfBlank(String value) {
    return ExcelDrawingBinarySupport.nullIfBlank(value);
  }

  String firstNonBlank(String first, String second) {
    return ExcelDrawingBinarySupport.firstNonBlank(first, second);
  }

  void updateAnchorInPlace(
      XSSFSheet sheet,
      String objectName,
      XmlObject parentAnchor,
      ExcelDrawingAnchor.TwoCell anchor) {
    ExcelDrawingAnchorSupport.updateAnchorInPlace(sheet, objectName, parentAnchor, anchor);
  }

  void removeTwoCellAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor target) {
    ExcelDrawingRemovalSupport.removeTwoCellAnchor(drawing, target);
  }

  void removeOneCellAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor target) {
    ExcelDrawingRemovalSupport.removeOneCellAnchor(drawing, target);
  }

  void removeAbsoluteAnchor(
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing drawing,
      org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor target) {
    ExcelDrawingRemovalSupport.removeAbsoluteAnchor(drawing, target);
  }

  org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs.Enum toPoiEditAs(
      ExcelDrawingAnchorBehavior behavior) {
    return ExcelDrawingAnchorSupport.toPoiEditAs(behavior);
  }

  private org.apache.poi.openxml4j.opc.PackagePart sheetPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ((XSSFSheet) objectData.getDrawing().getParent()).getPackagePart();
  }

  boolean looksLikeOle2Storage(byte[] bytes) throws IOException {
    return ExcelDrawingBinarySupport.looksLikeOle2Storage(bytes);
  }

  String partFileName(org.apache.poi.openxml4j.opc.PackagePart part) {
    return ExcelDrawingBinarySupport.partFileName(part);
  }

  record LocatedShape(
      XSSFDrawing drawing, XSSFShape shape, XmlObject shapeXml, XmlObject parentAnchor) {}

  record RasterDimensions(Integer widthPixels, Integer heightPixels) {
    static RasterDimensions none() {
      return new RasterDimensions(null, null);
    }
  }
}
