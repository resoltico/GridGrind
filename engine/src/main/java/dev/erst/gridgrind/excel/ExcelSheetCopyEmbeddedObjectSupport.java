package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingBinarySupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingRemovalSupport;
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
import org.jspecify.annotations.Nullable;

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
    for (EmbeddedObjectCopyPlan embeddedObject : snapshot.embeddedObjects) {
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
        ExcelDrawingBinarySupport.previewSheetRelationId(objectData.getOleObject()).orElse(null);
    @Nullable InternalRelationSnapshot previewSheetPart =
        previewSheetRelationId == null
            ? null
            : requiredInternalRelation(
                sourceSheet.getPackagePart(),
                previewSheetRelationId,
                objectName,
                "embedded object sheet preview");
    String previewDrawingRelationId =
        ExcelDrawingBinarySupport.previewDrawingRelationId(objectData).orElse(null);
    @Nullable InternalRelationSnapshot previewDrawingPart =
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
      @Nullable String relationshipId,
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
    PackagePart part =
        ExcelDrawingBinarySupport.relatedInternalPart(sourcePart, relationshipId).orElse(null);
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
        ExcelDrawingBinarySupport.previewSheetRelationId(targetObjectData.getOleObject())
            .orElse(null);
    if (embeddedObject.previewSheetPart() != null) {
      repairSheetPreviewRelation(
          targetSheet,
          targetObjectData,
          targetPreviewRelationId,
          embeddedObject.previewSheetPart(),
          embeddedObject.objectName());
    }
    String targetDrawingPreviewRelationId =
        ExcelDrawingBinarySupport.previewDrawingRelationId(targetObjectData).orElse(null);
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
        ExcelDrawingBinarySupport.blankAsOptional(targetObjectData.getOleObject().getId())
            .orElse(null);
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
      @Nullable String currentRelationId,
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
    @Nullable PackageRelationship existingRelationship = sheetPart.getRelationship(relationshipId);
    @Nullable PackagePart existingPart =
        ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, relationshipId).orElse(null);
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
      @Nullable PackageRelationship existingRelationship,
      @Nullable PackagePart existingPart,
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
      @Nullable String preferredRelationId,
      InternalRelationSnapshot sourcePart,
      String relationDescription,
      String objectName) {
    PackagePart sheetPart = targetSheet.getPackagePart();
    @Nullable PackageRelationship existingRelationship =
        preferredRelationId == null ? null : sheetPart.getRelationship(preferredRelationId);
    @Nullable PackagePart existingPart =
        preferredRelationId == null
            ? null
            : ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, preferredRelationId)
                .orElse(null);
    boolean referencedElsewhere =
        worksheetRelationIdReferencedElsewhere(
            targetSheet, targetObjectData, relationRole, preferredRelationId);
    if (matchingRelation(existingRelationship, existingPart, sourcePart) && !referencedElsewhere) {
      return Objects.requireNonNull(
          preferredRelationId, "preferredRelationId must not be null when relation matches");
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
      String resolvedPreferredRelationId =
          Objects.requireNonNull(preferredRelationId, "preferredRelationId must not be null");
      if (existingRelationship != null) {
        removeExistingRelationAndCleanup(sheetPart, resolvedPreferredRelationId, existingPart);
      }
      return sheetPart
          .addRelationship(
              copiedPart.getPartName(),
              TargetMode.INTERNAL,
              sourcePart.relationshipType(),
              resolvedPreferredRelationId)
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
      @Nullable PackagePart existingPart,
      InternalRelationSnapshot sourcePart) {
    return sourcePart.relationshipType().equals(existingRelationship.getRelationshipType())
        && (existingPart == null || sourcePart.contentType().equals(existingPart.getContentType()));
  }

  private static void removeExistingRelationAndCleanup(
      PackagePart sheetPart, String relationshipId, @Nullable PackagePart existingPart) {
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
      @Nullable String relationId) {
    String normalizedRelationId =
        ExcelDrawingBinarySupport.blankAsOptional(relationId).orElse(null);
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
        && relationId.equals(
            ExcelDrawingBinarySupport.blankAsOptional(oleObject.getId()).orElse(null));
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
        && relationId.equals(
            ExcelDrawingBinarySupport.previewSheetRelationId(oleObject).orElse(null));
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
        addNonBlank(
            referencedIds,
            ExcelDrawingBinarySupport.previewSheetRelationId(oleObject).orElse(null));
      }
    }
    return referencedIds;
  }

  private static void addNonBlank(Set<String> ids, @Nullable String value) {
    String normalized = ExcelDrawingBinarySupport.blankAsOptional(value).orElse(null);
    if (normalized != null) {
      ids.add(normalized);
    }
  }

  static void repairSheetDrawingRelation(XSSFSheet targetSheet) {
    if (!targetSheet.getCTWorksheet().isSetDrawing()) {
      return;
    }
    String drawingRelationId =
        ExcelDrawingBinarySupport.blankAsOptional(targetSheet.getCTWorksheet().getDrawing().getId())
            .orElse(null);
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
        ExcelDrawingBinarySupport.blankAsOptional(targetSheet.getCTWorksheet().getDrawing().getId())
            .orElse(null);
    if (drawingRelationId == null) {
      return;
    }
    PackagePart sheetPart = targetSheet.getPackagePart();
    PackageRelationship existingRelationship = sheetPart.getRelationship(drawingRelationId);
    PackagePart existingPart =
        ExcelDrawingBinarySupport.relatedInternalPart(sheetPart, drawingRelationId).orElse(null);
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

  /** Opaque embedded-object copy plan carried between sheet snapshot and repair phases. */
  static final class CopySnapshot {
    private final List<EmbeddedObjectCopyPlan> embeddedObjects;

    private CopySnapshot(List<EmbeddedObjectCopyPlan> embeddedObjects) {
      this.embeddedObjects = List.copyOf(embeddedObjects);
    }

    boolean isEmpty() {
      return embeddedObjects.isEmpty();
    }

    int embeddedObjectCount() {
      return embeddedObjects.size();
    }
  }

  record EmbeddedObjectCopyPlan(
      String objectName,
      InternalRelationSnapshot packagePart,
      @Nullable InternalRelationSnapshot previewSheetPart,
      @Nullable InternalRelationSnapshot previewDrawingPart) {
    EmbeddedObjectCopyPlan {
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
