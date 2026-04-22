package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Repairs embedded-object internals after POI sheet cloning leaves sheet-level relations behind.
 *
 * <p>POI clones the worksheet XML and drawing graph, but its clone path only recreates sheet
 * relations backed by POIXML document parts. Embedded-object package relationships live directly on
 * the sheet part, so cloned {@code oleObject} ids can point at nothing until GridGrind repairs
 * them.
 */
final class ExcelSheetCopyEmbeddedObjectSupport {
  CopySnapshot snapshot(ExcelSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    XSSFDrawing drawing = sourceSheet.xssfSheet().getDrawingPatriarch();
    if (drawing == null) {
      return new CopySnapshot(List.of());
    }

    List<EmbeddedObjectCopyPlan> embeddedObjects = new ArrayList<>();
    for (XSSFShape shape : drawing.getShapes()) {
      if (shape instanceof XSSFObjectData objectData) {
        embeddedObjects.add(snapshotEmbeddedObject(sourceSheet.xssfSheet(), objectData));
      }
    }
    return new CopySnapshot(embeddedObjects);
  }

  void repairCopiedEmbeddedObjects(ExcelSheet targetSheet, CopySnapshot snapshot) {
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    Objects.requireNonNull(snapshot, "snapshot must not be null");
    for (EmbeddedObjectCopyPlan embeddedObject : snapshot.embeddedObjects()) {
      repairCopiedEmbeddedObject(targetSheet.xssfSheet(), embeddedObject);
    }
  }

  private static EmbeddedObjectCopyPlan snapshotEmbeddedObject(
      XSSFSheet sourceSheet, XSSFObjectData objectData) {
    String objectName = ExcelDrawingAnchorSupport.resolvedName(objectData);
    String packageRelationId = objectData.getOleObject().getId();
    InternalRelationSnapshot packagePart =
        requiredInternalRelation(
            sourceSheet.getPackagePart(), packageRelationId, objectName, "embedded object package");

    String previewSheetRelationId =
        ExcelDrawingBinarySupport.previewSheetRelationId(objectData.getOleObject());
    InternalRelationSnapshot previewSheetPart =
        previewSheetRelationId == null
            ? null
            : requiredInternalRelation(
                sourceSheet.getPackagePart(),
                previewSheetRelationId,
                objectName,
                "embedded object sheet preview");
    return new EmbeddedObjectCopyPlan(objectName, packagePart, previewSheetPart);
  }

  static InternalRelationSnapshot requiredInternalRelation(
      PackagePart sourcePart,
      String relationshipId,
      String objectName,
      String relationDescription) {
    if (relationshipId == null || relationshipId.isBlank()) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its " + relationDescription + " id");
    }
    PackageRelationship relationship = sourcePart.getRelationship(relationshipId);
    if (relationship == null || relationship.getTargetMode() == TargetMode.EXTERNAL) {
      throw new IllegalStateException(
          "Embedded object '"
              + objectName
              + "' is missing its "
              + relationDescription
              + " relationship");
    }
    PackagePart part = ExcelDrawingBinarySupport.relatedInternalPart(sourcePart, relationshipId);
    if (part == null) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its " + relationDescription + " part");
    }
    return new InternalRelationSnapshot(
        relationship.getRelationshipType(),
        part.getContentType(),
        part.getPartName().getName(),
        ExcelBinaryData.readback(ExcelDrawingBinarySupport.partBytes(part)));
  }

  private static void repairCopiedEmbeddedObject(
      XSSFSheet targetSheet, EmbeddedObjectCopyPlan embeddedObject) {
    XSSFObjectData targetObjectData =
        requiredEmbeddedObject(targetSheet, embeddedObject.objectName());
    ensureInternalRelation(
        targetSheet.getPackagePart(),
        targetObjectData.getOleObject().getId(),
        embeddedObject.packagePart(),
        "embedded object package",
        embeddedObject.objectName());

    String targetPreviewRelationId =
        ExcelDrawingBinarySupport.previewSheetRelationId(targetObjectData.getOleObject());
    if (targetPreviewRelationId != null && embeddedObject.previewSheetPart() != null) {
      ensureInternalRelation(
          targetSheet.getPackagePart(),
          targetPreviewRelationId,
          embeddedObject.previewSheetPart(),
          "embedded object sheet preview",
          embeddedObject.objectName());
    }
  }

  static XSSFObjectData requiredEmbeddedObject(XSSFSheet targetSheet, String objectName) {
    XSSFDrawing drawing = targetSheet.getDrawingPatriarch();
    if (drawing == null) {
      throw new IllegalStateException(
          "Copied sheet '" + targetSheet.getSheetName() + "' is missing its drawing patriarch");
    }
    return drawing.getShapes().stream()
        .filter(XSSFObjectData.class::isInstance)
        .map(XSSFObjectData.class::cast)
        .filter(shape -> objectName.equals(ExcelDrawingAnchorSupport.resolvedName(shape)))
        .findFirst()
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "Copied sheet '"
                        + targetSheet.getSheetName()
                        + "' is missing embedded object '"
                        + objectName
                        + "'"));
  }

  static void ensureInternalRelation(
      PackagePart sheetPart,
      String relationshipId,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    PackagePart existingPart =
        ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, relationshipId);
    if (existingPart != null) {
      return;
    }
    if (sheetPart.getRelationship(relationshipId) != null) {
      sheetPart.removeRelationship(relationshipId);
    }
    PackagePart copiedPart =
        createCopiedPart(sheetPart.getPackage(), sourcePart, relationDescription, objectName);
    sheetPart.addRelationship(
        copiedPart.getPartName(),
        TargetMode.INTERNAL,
        sourcePart.relationshipType(),
        relationshipId);
  }

  static PackagePart createCopiedPart(
      OPCPackage pkg,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    try {
      PackagePartName targetPartName = nextCopiedPartName(pkg, sourcePart.sourcePartName());
      PackagePart copiedPart = pkg.createPart(targetPartName, sourcePart.contentType());
      try (var outputStream = copiedPart.getOutputStream()) {
        outputStream.write(sourcePart.bytes().bytes());
      }
      return copiedPart;
    } catch (IOException | InvalidFormatException exception) {
      throw new IllegalStateException(
          "Failed to copy " + relationDescription + " for embedded object '" + objectName + "'",
          exception);
    }
  }

  static PackagePartName nextCopiedPartName(OPCPackage pkg, String sourcePartName)
      throws InvalidFormatException {
    int extensionIndex = sourcePartName.lastIndexOf('.');
    String baseName =
        extensionIndex >= 0 ? sourcePartName.substring(0, extensionIndex) : sourcePartName;
    String extension = extensionIndex >= 0 ? sourcePartName.substring(extensionIndex) : "";
    for (int attempt = 1; ; attempt++) {
      PackagePartName candidate =
          PackagingURIHelper.createPartName(baseName + "-gridgrind-copy-" + attempt + extension);
      if (!pkg.containPart(candidate)) {
        return candidate;
      }
    }
  }

  record CopySnapshot(List<EmbeddedObjectCopyPlan> embeddedObjects) {
    CopySnapshot {
      embeddedObjects = List.copyOf(embeddedObjects);
    }
  }

  private record EmbeddedObjectCopyPlan(
      String objectName,
      InternalRelationSnapshot packagePart,
      InternalRelationSnapshot previewSheetPart) {
    private EmbeddedObjectCopyPlan {
      Objects.requireNonNull(objectName, "objectName must not be null");
      Objects.requireNonNull(packagePart, "packagePart must not be null");
    }
  }

  record InternalRelationSnapshot(
      String relationshipType, String contentType, String sourcePartName, ExcelBinaryData bytes) {
    InternalRelationSnapshot {
      relationshipType = requireNonBlank(relationshipType, "relationshipType");
      contentType = requireNonBlank(contentType, "contentType");
      sourcePartName = requireNonBlank(sourcePartName, "sourcePartName");
      Objects.requireNonNull(bytes, "bytes must not be null");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
