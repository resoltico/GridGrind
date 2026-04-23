package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
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
 * Repairs embedded-object internals after POI sheet cloning leaves sheet-level and preview-drawing
 * relations behind.
 *
 * <p>POI clones the worksheet XML and drawing graph, but its clone path only recreates some
 * relations backed by POIXML document parts. Embedded-object package relationships live directly on
 * the sheet part and embedded-object preview images can also be referenced from the drawing part,
 * so cloned {@code oleObject} ids or preview-image blips can point at nothing until GridGrind
 * repairs them.
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
    String previewDrawingRelationId =
        ExcelDrawingBinarySupport.previewDrawingRelationId(objectData);
    InternalRelationSnapshot previewDrawingPart =
        previewDrawingRelationId == null
            ? null
            : requiredInternalRelation(
                objectData.getDrawing().getPackagePart(),
                previewDrawingRelationId,
                objectName,
                "embedded object drawing preview");
    return new EmbeddedObjectCopyPlan(
        objectName, packagePart, previewSheetPart, previewDrawingPart);
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
    repairPackageRelation(
        targetSheet, targetObjectData, embeddedObject.packagePart(), embeddedObject.objectName());

    String targetPreviewRelationId =
        ExcelDrawingBinarySupport.previewSheetRelationId(targetObjectData.getOleObject());
    if (embeddedObject.previewSheetPart() != null) {
      repairSheetPreviewRelation(
          targetSheet,
          targetObjectData,
          targetPreviewRelationId,
          embeddedObject.previewSheetPart(),
          embeddedObject.objectName());
    }
    String targetDrawingPreviewRelationId =
        ExcelDrawingBinarySupport.previewDrawingRelationId(targetObjectData);
    if (targetDrawingPreviewRelationId != null && embeddedObject.previewDrawingPart() != null) {
      ensureInternalRelation(
          targetObjectData.getDrawing().getPackagePart(),
          targetDrawingPreviewRelationId,
          embeddedObject.previewDrawingPart(),
          "embedded object drawing preview",
          embeddedObject.objectName());
    }
    repairSheetDrawingRelation(targetSheet);
  }

  private static void repairPackageRelation(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      InternalRelationSnapshot sourcePart,
      String objectName) {
    String currentRelationId =
        ExcelDrawingBinarySupport.nullIfBlank(targetObjectData.getOleObject().getId());
    String repairedRelationId =
        repairWorksheetBoundRelation(
            targetSheet,
            targetObjectData,
            WorksheetRelationRole.OLE_OBJECT,
            currentRelationId,
            sourcePart,
            "embedded object package",
            objectName);
    targetObjectData.getOleObject().setId(repairedRelationId);
  }

  private static void repairSheetPreviewRelation(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      String currentRelationId,
      InternalRelationSnapshot sourcePart,
      String objectName) {
    String repairedRelationId =
        repairWorksheetBoundRelation(
            targetSheet,
            targetObjectData,
            WorksheetRelationRole.PREVIEW_SHEET,
            currentRelationId,
            sourcePart,
            "embedded object sheet preview",
            objectName);
    ExcelDrawingBinarySupport.setPreviewSheetRelationId(
        targetObjectData.getOleObject(), repairedRelationId);
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
    PackageRelationship existingRelationship = sheetPart.getRelationship(relationshipId);
    PackagePart existingPart =
        ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, relationshipId);
    if (matchingRelation(existingRelationship, existingPart, sourcePart)) {
      return;
    }
    if (existingRelationship != null) {
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

  private static boolean matchingRelation(
      PackageRelationship existingRelationship,
      PackagePart existingPart,
      InternalRelationSnapshot sourcePart) {
    if (existingRelationship == null || existingPart == null) {
      return false;
    }
    return sourcePart.relationshipType().equals(existingRelationship.getRelationshipType())
        && sourcePart.contentType().equals(existingPart.getContentType())
        && Arrays.equals(
            sourcePart.bytes().bytes(), ExcelDrawingBinarySupport.partBytes(existingPart));
  }

  static String repairWorksheetBoundRelation(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      WorksheetRelationRole relationRole,
      String preferredRelationId,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    PackagePart sheetPart = targetSheet.getPackagePart();
    PackageRelationship existingRelationship =
        preferredRelationId == null ? null : sheetPart.getRelationship(preferredRelationId);
    PackagePart existingPart =
        preferredRelationId == null
            ? null
            : ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, preferredRelationId);
    boolean referencedElsewhere =
        worksheetRelationIdReferencedElsewhere(
            targetSheet, targetObjectData, relationRole, preferredRelationId);
    if (matchingRelation(existingRelationship, existingPart, sourcePart) && !referencedElsewhere) {
      return preferredRelationId;
    }
    PackagePart copiedPart =
        createCopiedPart(sheetPart.getPackage(), sourcePart, relationDescription, objectName);
    boolean reusePreferredId =
        preferredRelationId != null
            && !referencedElsewhere
            && (existingRelationship == null
                || replaceableWorksheetBoundRelation(
                    existingRelationship, existingPart, sourcePart));
    if (reusePreferredId) {
      if (existingRelationship != null) {
        removeExistingRelationAndCleanup(sheetPart, preferredRelationId, existingPart);
      }
      return sheetPart
          .addRelationship(
              copiedPart.getPartName(),
              TargetMode.INTERNAL,
              sourcePart.relationshipType(),
              preferredRelationId)
          .getId();
    }
    return sheetPart
        .addRelationship(
            copiedPart.getPartName(),
            TargetMode.INTERNAL,
            sourcePart.relationshipType(),
            nextWorksheetRelationId(targetSheet))
        .getId();
  }

  private static boolean replaceableWorksheetBoundRelation(
      PackageRelationship existingRelationship,
      PackagePart existingPart,
      InternalRelationSnapshot sourcePart) {
    return sourcePart.relationshipType().equals(existingRelationship.getRelationshipType())
        && (existingPart == null || sourcePart.contentType().equals(existingPart.getContentType()));
  }

  private static void removeExistingRelationAndCleanup(
      PackagePart sheetPart, String relationshipId, PackagePart existingPart) {
    sheetPart.removeRelationship(relationshipId);
    if (existingPart != null) {
      ExcelDrawingRemovalSupport.cleanupPackagePartIfUnused(
          sheetPart.getPackage(), existingPart.getPartName());
    }
  }

  static boolean worksheetRelationIdReferencedElsewhere(
      XSSFSheet targetSheet,
      XSSFObjectData targetObjectData,
      WorksheetRelationRole relationRole,
      String relationId) {
    String normalizedRelationId = ExcelDrawingBinarySupport.nullIfBlank(relationId);
    if (normalizedRelationId == null) {
      return false;
    }
    var worksheet = targetSheet.getCTWorksheet();
    return worksheetStructureReferencesId(worksheet, normalizedRelationId)
        || (worksheet.isSetOleObjects()
            && oleObjectReferencesIdElsewhere(
                worksheet.getOleObjects().getOleObjectList(),
                targetObjectData.getOleObject(),
                relationRole,
                normalizedRelationId));
  }

  static boolean worksheetStructureReferencesId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet worksheet,
      String relationId) {
    return (worksheet.isSetDrawing() && relationId.equals(worksheet.getDrawing().getId()))
        || (worksheet.isSetLegacyDrawing()
            && relationId.equals(worksheet.getLegacyDrawing().getId()))
        || (worksheet.isSetLegacyDrawingHF()
            && relationId.equals(worksheet.getLegacyDrawingHF().getId()))
        || (worksheet.isSetDrawingHF() && relationId.equals(worksheet.getDrawingHF().getId()));
  }

  private static boolean oleObjectReferencesIdElsewhere(
      List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject> oleObjects,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject,
      WorksheetRelationRole relationRole,
      String relationId) {
    for (var oleObject : oleObjects) {
      if (referencesOleObjectRelationId(oleObject, targetOleObject, relationRole, relationId)
          || referencesPreviewRelationId(oleObject, targetOleObject, relationRole, relationId)) {
        return true;
      }
    }
    return false;
  }

  private static boolean referencesOleObjectRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject,
      WorksheetRelationRole relationRole,
      String relationId) {
    boolean differentTarget =
        relationRole != WorksheetRelationRole.OLE_OBJECT
            || !isTargetOleObject(oleObject, targetOleObject);
    return differentTarget
        && relationId.equals(ExcelDrawingBinarySupport.nullIfBlank(oleObject.getId()));
  }

  private static boolean referencesPreviewRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject,
      WorksheetRelationRole relationRole,
      String relationId) {
    boolean differentTarget =
        relationRole != WorksheetRelationRole.PREVIEW_SHEET
            || !isTargetOleObject(oleObject, targetOleObject);
    return differentTarget
        && relationId.equals(ExcelDrawingBinarySupport.previewSheetRelationId(oleObject));
  }

  private static boolean isTargetOleObject(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject targetOleObject) {
    return oleObject.getShapeId() == targetOleObject.getShapeId();
  }

  private static String nextWorksheetRelationId(XSSFSheet targetSheet) {
    return nextWorksheetRelationId(targetSheet, targetSheet.getPackagePart()::getRelationships);
  }

  static String nextWorksheetRelationId(
      XSSFSheet targetSheet, WorksheetRelationshipSupplier relationshipSupplier) {
    Set<String> reservedIds = referencedWorksheetRelationIds(targetSheet);
    try {
      for (PackageRelationship relationship : relationshipSupplier.relationships()) {
        reservedIds.add(relationship.getId());
      }
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(
          "Failed to inspect worksheet relationships for copied sheet '"
              + targetSheet.getSheetName()
              + "'",
          exception);
    }
    for (int attempt = 1; ; attempt++) {
      String candidate = "rId" + attempt;
      if (!reservedIds.contains(candidate)) {
        return candidate;
      }
    }
  }

  static Set<String> referencedWorksheetRelationIds(XSSFSheet targetSheet) {
    Set<String> referencedIds = new LinkedHashSet<>();
    var worksheet = targetSheet.getCTWorksheet();
    if (worksheet.isSetDrawing()) {
      addNonBlank(referencedIds, worksheet.getDrawing().getId());
    }
    if (worksheet.isSetLegacyDrawing()) {
      addNonBlank(referencedIds, worksheet.getLegacyDrawing().getId());
    }
    if (worksheet.isSetLegacyDrawingHF()) {
      addNonBlank(referencedIds, worksheet.getLegacyDrawingHF().getId());
    }
    if (worksheet.isSetDrawingHF()) {
      addNonBlank(referencedIds, worksheet.getDrawingHF().getId());
    }
    if (worksheet.isSetOleObjects()) {
      for (var oleObject : worksheet.getOleObjects().getOleObjectList()) {
        addNonBlank(referencedIds, oleObject.getId());
        addNonBlank(referencedIds, ExcelDrawingBinarySupport.previewSheetRelationId(oleObject));
      }
    }
    return referencedIds;
  }

  private static void addNonBlank(Set<String> ids, String value) {
    String normalized = ExcelDrawingBinarySupport.nullIfBlank(value);
    if (normalized != null) {
      ids.add(normalized);
    }
  }

  static void repairSheetDrawingRelation(XSSFSheet targetSheet) {
    if (!targetSheet.getCTWorksheet().isSetDrawing()) {
      return;
    }
    String drawingRelationId =
        ExcelDrawingBinarySupport.nullIfBlank(targetSheet.getCTWorksheet().getDrawing().getId());
    if (drawingRelationId == null) {
      return;
    }
    XSSFDrawing drawing = targetSheet.getDrawingPatriarch();
    if (drawing == null) {
      throw new IllegalStateException(
          "Copied sheet '" + targetSheet.getSheetName() + "' is missing its drawing patriarch");
    }
    repairSheetDrawingRelation(targetSheet, drawing);
  }

  static void repairSheetDrawingRelation(XSSFSheet targetSheet, XSSFDrawing drawing) {
    String drawingRelationId =
        ExcelDrawingBinarySupport.nullIfBlank(targetSheet.getCTWorksheet().getDrawing().getId());
    if (drawingRelationId == null) {
      return;
    }
    PackagePart sheetPart = targetSheet.getPackagePart();
    PackageRelationship existingRelationship = sheetPart.getRelationship(drawingRelationId);
    PackagePart existingPart =
        ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, drawingRelationId);
    if (existingRelationship != null
        && existingPart != null
        && org.apache.poi.xssf.usermodel.XSSFRelation.DRAWINGS
            .getRelation()
            .equals(existingRelationship.getRelationshipType())
        && drawing.getPackagePart().getPartName().equals(existingPart.getPartName())) {
      return;
    }
    if (existingRelationship != null) {
      sheetPart.removeRelationship(drawingRelationId);
    }
    sheetPart.addRelationship(
        drawing.getPackagePart().getPartName(),
        TargetMode.INTERNAL,
        org.apache.poi.xssf.usermodel.XSSFRelation.DRAWINGS.getRelation(),
        drawingRelationId);
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
      InternalRelationSnapshot previewSheetPart,
      InternalRelationSnapshot previewDrawingPart) {
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

  /**
   * Identifies which worksheet XML attribute currently owns one copied embedded-object relation.
   */
  enum WorksheetRelationRole {
    OLE_OBJECT,
    PREVIEW_SHEET
  }

  /** Supplies worksheet relationships for id allocation and failure-path coverage. */
  @FunctionalInterface
  interface WorksheetRelationshipSupplier {
    /** Returns the worksheet relationships that currently reserve relation ids. */
    Iterable<PackageRelationship> relationships() throws InvalidFormatException;
  }
}
