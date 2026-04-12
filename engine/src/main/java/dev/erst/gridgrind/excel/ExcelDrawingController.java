package dev.erst.gridgrind.excel;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Objects;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObjects;

/** Package-aware drawing and media controller for read and mutation workflows. */
final class ExcelDrawingController {
  private static final QName OLE_PREVIEW_ID =
      new QName(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS, "id");

  List<ExcelDrawingObjectSnapshot> drawingObjects(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) {
      return List.of();
    }

    List<ExcelDrawingObjectSnapshot> snapshots = new ArrayList<>();
    for (XSSFShape shape : drawing.getShapes()) {
      snapshots.add(snapshot(drawing, shape));
    }
    return List.copyOf(snapshots);
  }

  ExcelDrawingObjectPayload drawingObjectPayload(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");

    LocatedShape located = requiredLocatedShape(sheet, objectName);
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      return picturePayload(objectName, picture);
    }
    if (shape instanceof XSSFObjectData objectData) {
      return embeddedObjectPayload(objectName, objectData);
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' does not expose a binary payload");
  }

  void setPicture(XSSFSheet sheet, ExcelPictureDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedShapeIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    XSSFClientAnchor anchor = toPoiAnchor(drawing, definition.anchor());
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

  void setShape(XSSFSheet sheet, ExcelShapeDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedShapeIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    XSSFClientAnchor anchor = toPoiAnchor(drawing, definition.anchor());
    switch (definition.kind()) {
      case SIMPLE_SHAPE -> {
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.getCTShape().getNvSpPr().getCNvPr().setName(definition.name());
        shape.setShapeType(shapeType(definition.presetGeometryToken()));
        if (definition.text() != null) {
          shape.setText(definition.text());
        }
      }
      case CONNECTOR -> {
        XSSFConnector connector = drawing.createConnector(anchor);
        connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(definition.name());
      }
    }
  }

  void setEmbeddedObject(XSSFSheet sheet, ExcelEmbeddedObjectDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteNamedShapeIfPresent(sheet, definition.name());
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    XSSFClientAnchor anchor = toPoiAnchor(drawing, definition.anchor());
    XSSFObjectData objectData =
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

    LocatedShape located = requiredLocatedShape(sheet, objectName);
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture
        || shape instanceof XSSFObjectData
        || shape instanceof XSSFConnector
        || (shape instanceof XSSFSimpleShape && !(shape instanceof XSSFObjectData))) {
      updateAnchorInPlace(sheet, objectName, located.parentAnchor(), anchor);
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
    deleteLocatedShape(sheet, requiredLocatedShape(sheet, objectName));
  }

  void cleanupEmptyDrawingPatriarch(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    // POI exposes no supported API for unregistering a live XSSFDrawing relation object from its
    // parent sheet. Hard-deleting the package part leaves the in-memory relation graph stale and
    // breaks later save/commit flows on reopened workbooks. Prefer preserving an inert empty
    // drawing part over emitting a corrupt package.
  }

  private ExcelDrawingObjectSnapshot snapshot(XSSFDrawing drawing, XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return snapshotPicture(picture);
    }
    if (shape instanceof XSSFObjectData objectData) {
      return snapshotEmbeddedObject(objectData);
    }
    if (shape instanceof XSSFConnector connector) {
      return snapshotConnector(connector);
    }
    if (shape instanceof XSSFShapeGroup group) {
      return snapshotGroup(drawing, group);
    }
    if (shape instanceof XSSFGraphicFrame graphicFrame) {
      return snapshotGraphicFrame(graphicFrame);
    }
    if (shape instanceof XSSFSimpleShape simpleShape) {
      return snapshotShape(simpleShape);
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private ExcelDrawingObjectSnapshot.Picture snapshotPicture(XSSFPicture picture) {
    XSSFPictureData pictureData = picture.getPictureData();
    byte[] bytes = pictureData.getData();
    String description = nullIfBlank(picture.getCTPicture().getNvPicPr().getCNvPr().getDescr());
    return new ExcelDrawingObjectSnapshot.Picture(
        resolvedName(picture),
        snapshotAnchor(shapeXml(picture)),
        ExcelPictureFormat.fromPoiPictureType(pictureData.getPictureType()),
        pictureData.getPackagePart().getContentType(),
        bytes.length,
        sha256(bytes),
        null,
        null,
        description);
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotConnector(XSSFConnector connector) {
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(connector),
        snapshotAnchor(shapeXml(connector)),
        ExcelDrawingShapeKind.CONNECTOR,
        null,
        null,
        0);
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotGroup(
      XSSFDrawing drawing, XSSFShapeGroup group) {
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(group),
        snapshotAnchor(shapeXml(group)),
        ExcelDrawingShapeKind.GROUP,
        null,
        null,
        drawing.getShapes(group).size());
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotGraphicFrame(XSSFGraphicFrame graphicFrame) {
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(graphicFrame),
        snapshotAnchor(shapeXml(graphicFrame)),
        ExcelDrawingShapeKind.GRAPHIC_FRAME,
        null,
        null,
        0);
  }

  private ExcelDrawingObjectSnapshot.Shape snapshotShape(XSSFSimpleShape shape) {
    STShapeType.Enum preset =
        shape.getCTShape().getSpPr() == null || !shape.getCTShape().getSpPr().isSetPrstGeom()
            ? null
            : shape.getCTShape().getSpPr().getPrstGeom().getPrst();
    return new ExcelDrawingObjectSnapshot.Shape(
        resolvedName(shape),
        snapshotAnchor(shapeXml(shape)),
        ExcelDrawingShapeKind.SIMPLE_SHAPE,
        preset == null ? null : preset.toString(),
        nullIfBlank(shape.getText()),
        0);
  }

  private ExcelDrawingObjectSnapshot.EmbeddedObject snapshotEmbeddedObject(
      XSSFObjectData objectData) {
    EmbeddedObjectReadback readback = embeddedObjectReadback(resolvedName(objectData), objectData);
    return new ExcelDrawingObjectSnapshot.EmbeddedObject(
        resolvedName(objectData),
        snapshotAnchor(shapeXml(objectData)),
        readback.packagingKind(),
        readback.label(),
        readback.fileName(),
        readback.command(),
        readback.contentType(),
        readback.payload().size(),
        sha256(readback.payload().bytes()),
        readback.previewFormat(),
        readback.previewImage() == null ? null : (long) readback.previewImage().size(),
        readback.previewImage() == null ? null : sha256(readback.previewImage().bytes()));
  }

  private ExcelDrawingObjectPayload.Picture picturePayload(String objectName, XSSFPicture picture) {
    XSSFPictureData pictureData = picture.getPictureData();
    byte[] bytes = pictureData.getData();
    ExcelPictureFormat format = ExcelPictureFormat.fromPoiPictureType(pictureData.getPictureType());
    return new ExcelDrawingObjectPayload.Picture(
        objectName,
        format,
        pictureData.getPackagePart().getContentType(),
        objectName + format.defaultExtension(),
        sha256(bytes),
        new ExcelBinaryData(bytes),
        nullIfBlank(picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()));
  }

  private ExcelDrawingObjectPayload.EmbeddedObject embeddedObjectPayload(
      String objectName, XSSFObjectData objectData) {
    EmbeddedObjectReadback readback = embeddedObjectReadback(objectName, objectData);
    return new ExcelDrawingObjectPayload.EmbeddedObject(
        objectName,
        readback.packagingKind(),
        readback.contentType(),
        readback.fileName(),
        sha256(readback.payload().bytes()),
        readback.payload(),
        readback.label(),
        readback.command());
  }

  private EmbeddedObjectReadback embeddedObjectReadback(
      String objectName, XSSFObjectData objectData) {
    PackagePart previewPart = previewImagePart(objectData);
    ExcelBinaryData previewImage =
        previewPart == null ? null : binary(partBytes(previewPart), "preview image");
    ExcelPictureFormat previewFormat =
        previewPart == null
            ? null
            : ExcelPictureFormat.fromContentType(previewPart.getContentType());
    PackagePart objectPart = oleObjectPart(objectData);
    if (objectPart == null) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its package relationship");
    }
    ExcelBinaryData rawPackage;
    rawPackage = binary(partBytes(objectPart), "embedded object package");
    String contentType = objectPart.getContentType();
    String label = null;
    String fileName = partFileName(objectPart);
    String command = null;
    ExcelEmbeddedObjectPackagingKind packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    ExcelBinaryData payload = rawPackage;
    try {
      if (looksLikeOle2Storage(rawPackage.bytes())) {
        Ole10Native nativeData = ole10Native(rawPackage.bytes());
        payload = new ExcelBinaryData(nativeData.getDataBuffer());
        label = firstNonBlank(nativeData.getLabel2(), nativeData.getLabel());
        fileName = firstNonBlank(nativeData.getFileName2(), nativeData.getFileName());
        command = firstNonBlank(nativeData.getCommand2(), nativeData.getCommand());
        packagingKind = ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE;
        contentType = "application/octet-stream";
      }
    } catch (IOException | Ole10NativeException exception) {
      // Preserve truthful raw-package readback when the embedded object is not a POI-packaged
      // Ole10Native payload.
      payload = rawPackage;
      packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    }
    fileName = Objects.requireNonNullElse(fileName, objectName + ".bin");
    return new EmbeddedObjectReadback(
        packagingKind, label, fileName, command, contentType, payload, previewFormat, previewImage);
  }

  private void deleteNamedShapeIfPresent(XSSFSheet sheet, String objectName) {
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    if (located != null) {
      deleteLocatedShape(sheet, located);
    }
  }

  private void deleteLocatedShape(XSSFSheet sheet, LocatedShape located) {
    XSSFShape shape = located.shape();
    if (shape instanceof XSSFPicture picture) {
      deletePicture(sheet, located, picture);
      return;
    }
    if (shape instanceof XSSFObjectData objectData) {
      deleteEmbeddedObject(sheet, located, objectData);
      return;
    }
    if (shape instanceof XSSFConnector
        || (shape instanceof XSSFSimpleShape && !(shape instanceof XSSFObjectData))
        || shape instanceof XSSFShapeGroup) {
      removeParentAnchor(located.drawing(), located.parentAnchor());
      cleanupEmptyDrawingPatriarch(sheet);
      return;
    }
    throw new IllegalArgumentException(
        "Drawing object '"
            + resolvedName(located.shape())
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is read-only until a later parity phase");
  }

  private void deletePicture(XSSFSheet sheet, LocatedShape located, XSSFPicture picture) {
    PackagePartName imagePartName = picture.getPictureData().getPackagePart().getPartName();
    String relationId = picture.getCTPicture().getBlipFill().getBlip().getEmbed();
    located.drawing().getPackagePart().removeRelationship(relationId);
    removeParentAnchor(located.drawing(), located.parentAnchor());
    cleanupWorkbookImagePartIfUnused(sheet.getWorkbook(), imagePartName);
    cleanupEmptyDrawingPatriarch(sheet);
  }

  private void deleteEmbeddedObject(
      XSSFSheet sheet, LocatedShape located, XSSFObjectData objectData) {
    PackagePart objectPart = oleObjectPart(objectData);
    PackagePart previewPart = previewImagePart(objectData);
    PackagePartName olePartName = objectPart == null ? null : objectPart.getPartName();
    PackagePartName previewPartName = previewPart == null ? null : previewPart.getPartName();
    String drawingPreviewRelationId = previewDrawingRelationId(objectData);
    String sheetPreviewRelationId = previewSheetRelationId(objectData.getOleObject());
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
    removeOleObject(sheet, objectData.getOleObject());
    removeParentAnchor(located.drawing(), located.parentAnchor());
    cleanupPackagePartIfUnused(sheet.getWorkbook().getPackage(), olePartName);
    cleanupWorkbookImagePartIfUnused(sheet.getWorkbook(), previewPartName);
    cleanupEmptyDrawingPatriarch(sheet);
  }

  private PackagePart oleObjectPart(XSSFObjectData objectData) {
    return relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId());
  }

  private PackagePart previewImagePart(XSSFObjectData objectData) {
    String sheetRelationId = previewSheetRelationId(objectData.getOleObject());
    if (sheetRelationId != null) {
      PackagePart previewPart = relatedInternalPart(sheetPart(objectData), sheetRelationId);
      if (previewPart != null) {
        return previewPart;
      }
    }
    String drawingRelationId = previewDrawingRelationId(objectData);
    return drawingRelationId == null
        ? null
        : relatedInternalPart(objectData.getDrawing().getPackagePart(), drawingRelationId);
  }

  private void removeOleObject(XSSFSheet sheet, CTOleObject target) {
    CTOleObjects oleObjects = sheet.getCTWorksheet().getOleObjects();
    if (oleObjects == null) {
      return;
    }
    for (int index = 0; index < oleObjects.sizeOfOleObjectArray(); index++) {
      CTOleObject candidate = oleObjects.getOleObjectArray(index);
      if (candidate.equals(target)) {
        oleObjects.removeOleObject(index);
        if (oleObjects.sizeOfOleObjectArray() == 0) {
          sheet.getCTWorksheet().unsetOleObjects();
        }
        return;
      }
    }
  }

  private void cleanupWorkbookImagePartIfUnused(
      XSSFWorkbook workbook, PackagePartName imagePartName) {
    if (imagePartName == null || imagePartUsed(workbook, imagePartName)) {
      return;
    }
    removeRelationshipsToPart(workbook.getPackagePart(), imagePartName);
    cleanupPackagePartIfUnused(workbook.getPackage(), imagePartName);
  }

  private boolean imagePartUsed(XSSFWorkbook workbook, PackagePartName imagePartName) {
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      XSSFDrawing drawing = workbook.getSheetAt(sheetIndex).getDrawingPatriarch();
      if (drawing == null) {
        continue;
      }
      for (XSSFShape shape : drawing.getShapes()) {
        if (shape instanceof XSSFPicture picture
            && picture.getPictureData().getPackagePart().getPartName().equals(imagePartName)) {
          return true;
        }
        if (shape instanceof XSSFObjectData objectData) {
          PackagePart previewPart = previewImagePart(objectData);
          if (previewPart != null && previewPart.getPartName().equals(imagePartName)) {
            return true;
          }
        }
      }
    }
    return false;
  }

  private void removeRelationshipsToPart(PackagePart sourcePart, PackagePartName targetPartName) {
    if (sourcePart.isRelationshipPart()) {
      return;
    }
    List<String> relationshipIds = new ArrayList<>();
    try {
      for (PackageRelationship relationship : sourcePart.getRelationships()) {
        if (relationship.getTargetMode() == TargetMode.EXTERNAL) {
          continue;
        }
        if (targetPartName
            .getURI()
            .equals(
                PackagingURIHelper.resolvePartUri(
                    sourcePart.getPartName().getURI(), relationship.getTargetURI()))) {
          relationshipIds.add(relationship.getId());
        }
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect package relationships", exception);
    }
    for (String relationshipId : relationshipIds) {
      sourcePart.removeRelationship(relationshipId);
    }
  }

  private PackagePart relatedInternalPart(PackagePart sourcePart, String relationshipId) {
    if (relationshipId == null || relationshipId.isBlank()) {
      return null;
    }
    PackageRelationship relationship = sourcePart.getRelationship(relationshipId);
    if (relationship == null || relationship.getTargetMode() == TargetMode.EXTERNAL) {
      return null;
    }
    try {
      return sourcePart
          .getPackage()
          .getPart(
              PackagingURIHelper.createPartName(
                  PackagingURIHelper.resolvePartUri(
                      sourcePart.getPartName().getURI(), relationship.getTargetURI())));
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to resolve related package part", exception);
    }
  }

  private void cleanupPackagePartIfUnused(OPCPackage pkg, PackagePartName partName) {
    if (partName == null) {
      return;
    }
    try {
      for (PackagePart part : pkg.getParts()) {
        if (part.isRelationshipPart()) {
          continue;
        }
        for (PackageRelationship relationship : part.getRelationships()) {
          if (relationship.getTargetMode() == TargetMode.EXTERNAL) {
            continue;
          }
          if (partName
              .getURI()
              .equals(
                  PackagingURIHelper.resolvePartUri(
                      part.getPartName().getURI(), relationship.getTargetURI()))) {
            return;
          }
        }
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect package relationships", exception);
    }
    if (pkg.containPart(partName)) {
      pkg.deletePartRecursive(partName);
    }
  }

  private void removeParentAnchor(XSSFDrawing drawing, XmlObject parentAnchor) {
    switch (parentAnchor) {
      case CTTwoCellAnchor twoCellAnchor ->
          removeTwoCellAnchor(drawing.getCTDrawing(), twoCellAnchor);
      case CTOneCellAnchor oneCellAnchor ->
          removeOneCellAnchor(drawing.getCTDrawing(), oneCellAnchor);
      case CTAbsoluteAnchor absoluteAnchor ->
          removeAbsoluteAnchor(drawing.getCTDrawing(), absoluteAnchor);
      case null -> throw new IllegalStateException("Drawing object is missing its parent anchor");
      default ->
          throw new IllegalStateException(
              "Unsupported parent anchor type: " + parentAnchor.getClass().getName());
    }
  }

  private void removeTwoCellAnchor(CTDrawing drawing, CTTwoCellAnchor target) {
    for (int index = 0; index < drawing.sizeOfTwoCellAnchorArray(); index++) {
      CTTwoCellAnchor candidate = drawing.getTwoCellAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeTwoCellAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate two-cell drawing anchor for deletion");
  }

  private void removeOneCellAnchor(CTDrawing drawing, CTOneCellAnchor target) {
    for (int index = 0; index < drawing.sizeOfOneCellAnchorArray(); index++) {
      CTOneCellAnchor candidate = drawing.getOneCellAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeOneCellAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate one-cell drawing anchor for deletion");
  }

  private void removeAbsoluteAnchor(CTDrawing drawing, CTAbsoluteAnchor target) {
    for (int index = 0; index < drawing.sizeOfAbsoluteAnchorArray(); index++) {
      CTAbsoluteAnchor candidate = drawing.getAbsoluteAnchorArray(index);
      if (candidate.equals(target)) {
        drawing.removeAbsoluteAnchor(index);
        return;
      }
    }
    throw new IllegalStateException("Failed to locate absolute drawing anchor for deletion");
  }

  private LocatedShape requiredLocatedShape(XSSFSheet sheet, String objectName) {
    LocatedShape located = optionalLocatedShape(sheet, objectName);
    if (located == null) {
      throw new DrawingObjectNotFoundException(sheet.getSheetName(), objectName);
    }
    return located;
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

  private XSSFClientAnchor toPoiAnchor(XSSFDrawing drawing, ExcelDrawingAnchor.TwoCell anchor) {
    XSSFClientAnchor poiAnchor =
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

  private ClientAnchor.AnchorType toPoiBehavior(ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE -> ClientAnchor.AnchorType.MOVE_AND_RESIZE;
      case MOVE_DONT_RESIZE -> ClientAnchor.AnchorType.MOVE_DONT_RESIZE;
      case DONT_MOVE_AND_RESIZE -> ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
    };
  }

  private ExcelDrawingAnchor snapshotAnchor(XmlObject shapeXml) {
    XmlObject parentAnchor = parentAnchor(shapeXml);
    return switch (parentAnchor) {
      case CTTwoCellAnchor twoCellAnchor ->
          new ExcelDrawingAnchor.TwoCell(
              marker(twoCellAnchor.getFrom()),
              marker(twoCellAnchor.getTo()),
              behavior(twoCellAnchor.getEditAs()));
      case CTOneCellAnchor oneCellAnchor ->
          new ExcelDrawingAnchor.OneCell(
              marker(oneCellAnchor.getFrom()),
              oneCellAnchor.getExt().getCx(),
              oneCellAnchor.getExt().getCy(),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      case CTAbsoluteAnchor absoluteAnchor ->
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

  private ExcelDrawingAnchorBehavior behavior(STEditAs.Enum editAs) {
    return switch (editAs == null ? STEditAs.INT_TWO_CELL : editAs.intValue()) {
      case STEditAs.INT_ONE_CELL -> ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE;
      case STEditAs.INT_ABSOLUTE -> ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE;
      default -> ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE;
    };
  }

  private ExcelDrawingMarker marker(CTMarker marker) {
    return new ExcelDrawingMarker(
        marker.getCol(),
        marker.getRow(),
        Math.toIntExact(((Number) marker.getColOff()).longValue()),
        Math.toIntExact(((Number) marker.getRowOff()).longValue()));
  }

  private String previewSheetRelationId(CTOleObject oleObject) {
    try (XmlCursor cursor = oleObject.newCursor()) {
      if (cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr")) {
        return cursor.getAttributeText(OLE_PREVIEW_ID);
      }
    }
    return null;
  }

  private String previewDrawingRelationId(XSSFObjectData objectData) {
    if (objectData.getCTShape().getSpPr() == null
        || objectData.getCTShape().getSpPr().getBlipFill() == null
        || objectData.getCTShape().getSpPr().getBlipFill().getBlip() == null) {
      return null;
    }
    return objectData.getCTShape().getSpPr().getBlipFill().getBlip().getEmbed();
  }

  private XmlObject shapeXml(XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return picture.getCTPicture();
    }
    if (shape instanceof XSSFObjectData objectData) {
      return objectData.getCTShape();
    }
    if (shape instanceof XSSFConnector connector) {
      return connector.getCTConnector();
    }
    if (shape instanceof XSSFShapeGroup group) {
      return group.getCTGroupShape();
    }
    if (shape instanceof XSSFGraphicFrame graphicFrame) {
      return graphicFrame.getCTGraphicalObjectFrame();
    }
    if (shape instanceof XSSFSimpleShape simpleShape) {
      return simpleShape.getCTShape();
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private XmlObject parentAnchor(XmlObject shapeXml) {
    try (XmlCursor cursor = shapeXml.newCursor()) {
      if (cursor.toParent()) {
        return cursor.getObject();
      }
    }
    return null;
  }

  private String resolvedName(XSSFShape shape) {
    return nullIfBlank(shape.getShapeName()) != null ? shape.getShapeName() : defaultName(shape);
  }

  private String defaultName(XSSFShape shape) {
    if (shape instanceof XSSFPicture picture) {
      return "Picture-" + picture.getCTPicture().getNvPicPr().getCNvPr().getId();
    }
    if (shape instanceof XSSFObjectData objectData) {
      return "Object-" + objectData.getCTShape().getNvSpPr().getCNvPr().getId();
    }
    if (shape instanceof XSSFConnector connector) {
      return "Connector-" + connector.getCTConnector().getNvCxnSpPr().getCNvPr().getId();
    }
    if (shape instanceof XSSFShapeGroup group) {
      return "Group-" + group.getCTGroupShape().getNvGrpSpPr().getCNvPr().getId();
    }
    if (shape instanceof XSSFGraphicFrame graphicFrame) {
      return "GraphicFrame-"
          + graphicFrame.getCTGraphicalObjectFrame().getNvGraphicFramePr().getCNvPr().getId();
    }
    if (shape instanceof XSSFSimpleShape simpleShape) {
      return "Shape-" + simpleShape.getCTShape().getNvSpPr().getCNvPr().getId();
    }
    throw new IllegalStateException(
        "Unsupported drawing shape type: " + shape.getClass().getName());
  }

  private int shapeType(String presetGeometryToken) {
    STShapeType.Enum preset = STShapeType.Enum.forString(presetGeometryToken);
    if (preset == null) {
      throw new IllegalArgumentException("Unsupported presetGeometryToken: " + presetGeometryToken);
    }
    return preset.intValue();
  }

  private ExcelBinaryData binary(byte[] bytes, String label) {
    try {
      return new ExcelBinaryData(bytes);
    } catch (IllegalArgumentException exception) {
      throw new IllegalStateException("Missing " + label + " bytes", exception);
    }
  }

  private byte[] partBytes(PackagePart part) {
    try {
      return part.getInputStream().readAllBytes();
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read package part bytes", exception);
    }
  }

  private String sha256(byte[] bytes) {
    try {
      MessageDigest digest = MessageDigest.getInstance("SHA-256");
      return HexFormat.of().formatHex(digest.digest(bytes));
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 digest is unavailable", exception);
    }
  }

  private String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }

  private String firstNonBlank(String first, String second) {
    String normalizedFirst = nullIfBlank(first);
    return normalizedFirst != null ? normalizedFirst : nullIfBlank(second);
  }

  private void updateAnchorInPlace(
      XSSFSheet sheet,
      String objectName,
      XmlObject parentAnchor,
      ExcelDrawingAnchor.TwoCell anchor) {
    if (!(parentAnchor instanceof CTTwoCellAnchor twoCellAnchor)) {
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

  private void applyMarker(CTMarker target, ExcelDrawingMarker source) {
    target.setCol(source.columnIndex());
    target.setRow(source.rowIndex());
    target.setColOff(source.dx());
    target.setRowOff(source.dy());
  }

  private STEditAs.Enum toPoiEditAs(ExcelDrawingAnchorBehavior behavior) {
    return switch (behavior) {
      case MOVE_AND_RESIZE -> STEditAs.TWO_CELL;
      case MOVE_DONT_RESIZE -> STEditAs.ONE_CELL;
      case DONT_MOVE_AND_RESIZE -> STEditAs.ABSOLUTE;
    };
  }

  private PackagePart sheetPart(XSSFObjectData objectData) {
    return ((XSSFSheet) objectData.getDrawing().getParent()).getPackagePart();
  }

  private boolean looksLikeOle2Storage(byte[] bytes) throws IOException {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      return FileMagic.valueOf(FileMagic.prepareToCheckMagic(input)) == FileMagic.OLE2;
    }
  }

  private Ole10Native ole10Native(byte[] bytes) throws IOException, Ole10NativeException {
    try (POIFSFileSystem filesystem = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
      DirectoryNode directory = filesystem.getRoot();
      return Ole10Native.createFromEmbeddedOleObject(directory);
    }
  }

  private String partFileName(PackagePart part) {
    String partName = part.getPartName().getName();
    return partName.substring(partName.lastIndexOf('/') + 1);
  }

  private record LocatedShape(
      XSSFDrawing drawing, XSSFShape shape, XmlObject shapeXml, XmlObject parentAnchor) {}

  private record EmbeddedObjectReadback(
      ExcelEmbeddedObjectPackagingKind packagingKind,
      String label,
      String fileName,
      String command,
      String contentType,
      ExcelBinaryData payload,
      ExcelPictureFormat previewFormat,
      ExcelBinaryData previewImage) {}
}
