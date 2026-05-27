package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelPackageRelationshipSupport;
import dev.erst.gridgrind.excel.ExcelWorkbookImageCatalogSupport;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;

/** Binary drawing payload, preview-image, and embedded-object package helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelDrawingBinarySupport {
  private ExcelDrawingBinarySupport() {}

  public static ExcelDrawingObjectSnapshot.EmbeddedObject snapshotEmbeddedObject(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    String objectName = ExcelDrawingAnchorSupport.resolvedName(objectData);
    EmbeddedObjectReadback readback = embeddedObjectReadback(objectName, objectData);
    return new ExcelDrawingObjectSnapshot.EmbeddedObject(
        objectName,
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(objectData)),
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

  public static ExcelDrawingObjectPayload.Picture picturePayload(
      String objectName, XSSFPicture picture) {
    return ExcelDrawingPictureSupport.picturePayload(objectName, picture);
  }

  public static ExcelDrawingObjectPayload.EmbeddedObject embeddedObjectPayload(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
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

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> previewImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart =
        previewSheetImagePart(objectData);
    if (previewPart.isPresent()) {
      return previewPart;
    }
    Optional<String> drawingRelationId = previewDrawingRelationId(objectData);
    Optional<org.apache.poi.openxml4j.opc.PackagePart> drawingPreviewPart =
        drawingRelationId.flatMap(
            relationId ->
                relatedInternalPart(objectData.getDrawing().getPackagePart(), relationId));
    return drawingPreviewPart.filter(ExcelDrawingBinarySupport::supportedPreviewImagePart);
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> previewSheetImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    Optional<String> sheetRelationId = previewSheetRelationId(objectData.getOleObject());
    if (sheetRelationId.isEmpty()) {
      return Optional.empty();
    }
    Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart =
        relatedInternalPart(sheetPart(objectData), sheetRelationId.orElseThrow());
    return previewPart.filter(ExcelDrawingBinarySupport::supportedPreviewImagePart);
  }

  public static Optional<String> previewSheetImageRelationId(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return previewImagePart(objectData).isPresent()
        ? previewSheetRelationId(objectData.getOleObject())
        : Optional.empty();
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> relatedInternalPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart, @Nullable String relationshipId) {
    if (relationshipId == null || relationshipId.isBlank()) {
      return Optional.empty();
    }
    org.apache.poi.openxml4j.opc.PackageRelationship relationship =
        sourcePart.getRelationship(relationshipId);
    if (relationship == null
        || relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
      return Optional.empty();
    }
    try {
      return Optional.ofNullable(
          sourcePart
              .getPackage()
              .getPart(
                  org.apache.poi.openxml4j.opc.PackagingURIHelper.createPartName(
                      org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                          sourcePart.getPartName().getURI(), relationship.getTargetURI()))));
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to resolve related package part", exception);
    }
  }

  public static void cleanupWorkbookImagePartIfUnused(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.@Nullable PackagePartName imagePartName) {
    if (imagePartName == null) {
      return;
    }
    if (!imagePartUsed(workbook, imagePartName)) {
      removeRelationshipsToPart(workbook.getPackagePart(), imagePartName);
      cleanupPackagePartIfUnused(workbook.getPackage(), imagePartName);
    }
    ExcelWorkbookImageCatalogSupport.synchronizePictureCatalog(workbook);
  }

  public static void removeOleObject(
      XSSFSheet sheet, org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject target) {
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObjects oleObjects =
        sheet.getCTWorksheet().getOleObjects();
    if (oleObjects == null) {
      return;
    }
    for (int index = 0; index < oleObjects.sizeOfOleObjectArray(); index++) {
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject candidate =
          oleObjects.getOleObjectArray(index);
      if (candidate.equals(target)) {
        oleObjects.removeOleObject(index);
        if (oleObjects.sizeOfOleObjectArray() == 0) {
          sheet.getCTWorksheet().unsetOleObjects();
        }
        return;
      }
    }
  }

  public static Optional<String> previewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject) {
    if (!oleObject.isSetObjectPr()) {
      return Optional.empty();
    }
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTObjectPr objectPr =
        oleObject.getObjectPr();
    return objectPr.isSetId() ? blankAsOptional(objectPr.getId()) : Optional.empty();
  }

  public static void setPreviewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject,
      String relationId) {
    Objects.requireNonNull(oleObject, "oleObject must not be null");
    Objects.requireNonNull(relationId, "relationId must not be null");
    if (relationId.isBlank()) {
      throw new IllegalArgumentException("relationId must not be blank");
    }
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTObjectPr objectPr =
        oleObject.isSetObjectPr() ? oleObject.getObjectPr() : oleObject.addNewObjectPr();
    objectPr.setId(relationId);
  }

  public static Optional<String> previewDrawingRelationId(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    if (objectData.getCTShape().getSpPr().getBlipFill() == null
        || objectData.getCTShape().getSpPr().getBlipFill().getBlip() == null) {
      return Optional.empty();
    }
    return blankAsOptional(objectData.getCTShape().getSpPr().getBlipFill().getBlip().getEmbed());
  }

  public static ExcelBinaryData binary(byte[] bytes, String label) {
    Objects.requireNonNull(label, "label must not be null");
    return ExcelBinaryData.readback(bytes);
  }

  public static byte[] partBytes(org.apache.poi.openxml4j.opc.PackagePart part) {
    try {
      return part.getInputStream().readAllBytes();
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read package part bytes", exception);
    }
  }

  public static String sha256(byte[] bytes) {
    try {
      MessageDigest digest = MessageDigest.getInstance("SHA-256");
      return HexFormat.of().formatHex(digest.digest(bytes));
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 digest is unavailable", exception);
    }
  }

  public static Optional<String> blankAsOptional(@Nullable String value) {
    return value == null || value.isBlank() ? Optional.empty() : Optional.of(value);
  }

  public static Optional<String> firstNonBlank(String first, String second) {
    Optional<String> normalizedFirst = blankAsOptional(first);
    return normalizedFirst.isPresent() ? normalizedFirst : blankAsOptional(second);
  }

  public static boolean looksLikeOle2Storage(byte[] bytes) throws IOException {
    if (bytes.length == 0) {
      return false;
    }
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      return org.apache.poi.poifs.filesystem.FileMagic.valueOf(
              org.apache.poi.poifs.filesystem.FileMagic.prepareToCheckMagic(input))
          == org.apache.poi.poifs.filesystem.FileMagic.OLE2;
    }
  }

  public static String partFileName(org.apache.poi.openxml4j.opc.PackagePart part) {
    String partName = part.getPartName().getName();
    return partName.substring(partName.lastIndexOf('/') + 1);
  }

  public static boolean supportedPreviewImagePart(org.apache.poi.openxml4j.opc.PackagePart part) {
    if (part == null) {
      return false;
    }
    try {
      ExcelPictureFormat.fromContentType(part.getContentType());
      return true;
    } catch (IllegalArgumentException ignored) {
      return false;
    }
  }

  private static EmbeddedObjectReadback embeddedObjectReadback(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart = previewImagePart(objectData);
    ExcelBinaryData previewImage =
        previewPart.isPresent()
            ? binary(partBytes(previewPart.orElseThrow()), "preview image")
            : null;
    ExcelPictureFormat previewFormat =
        previewPart.isPresent()
            ? ExcelPictureFormat.fromContentType(previewPart.orElseThrow().getContentType())
            : null;
    Optional<org.apache.poi.openxml4j.opc.PackagePart> objectPart = oleObjectPart(objectData);
    if (objectPart.isEmpty()) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its package relationship");
    }
    org.apache.poi.openxml4j.opc.PackagePart resolvedObjectPart = objectPart.orElseThrow();
    ExcelBinaryData rawPackage = binary(partBytes(resolvedObjectPart), "embedded object package");
    String contentType = resolvedObjectPart.getContentType();
    String label = null;
    String fileName = partFileName(resolvedObjectPart);
    String command = null;
    ExcelEmbeddedObjectPackagingKind packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    ExcelBinaryData payload = rawPackage;
    try {
      if (looksLikeOle2Storage(rawPackage.bytes())) {
        org.apache.poi.poifs.filesystem.Ole10Native nativeData = ole10Native(rawPackage.bytes());
        payload = ExcelBinaryData.readback(nativeData.getDataBuffer());
        label = firstNonBlank(nativeData.getLabel2(), nativeData.getLabel()).orElse(null);
        fileName = firstNonBlank(nativeData.getFileName2(), nativeData.getFileName()).orElse(null);
        command = firstNonBlank(nativeData.getCommand2(), nativeData.getCommand()).orElse(null);
        packagingKind = ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE;
        contentType = "application/octet-stream";
      }
    } catch (IOException | org.apache.poi.poifs.filesystem.Ole10NativeException exception) {
      payload = rawPackage;
      packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    }
    fileName = Objects.requireNonNullElse(fileName, objectName + ".bin");
    return new EmbeddedObjectReadback(
        packagingKind, label, fileName, command, contentType, payload, previewFormat, previewImage);
  }

  public static Optional<org.apache.poi.openxml4j.opc.PackagePart> oleObjectPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId());
  }

  public static boolean imagePartUsed(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    ExcelSignatureLineController signatureLineController = new ExcelSignatureLineController();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      if (sheetUsesImagePart(
          workbook.getSheetAt(sheetIndex), imagePartName, signatureLineController)) {
        return true;
      }
    }
    return false;
  }

  private static boolean sheetUsesImagePart(
      XSSFSheet sheet,
      org.apache.poi.openxml4j.opc.PackagePartName imagePartName,
      ExcelSignatureLineController signatureLineController) {
    org.apache.poi.xssf.usermodel.XSSFDrawing drawing = sheet.getDrawingPatriarch();
    return (drawing != null && drawingUsesImagePart(drawing, imagePartName))
        || signatureLineController.usesImagePart(sheet, imagePartName);
  }

  private static boolean drawingUsesImagePart(
      org.apache.poi.xssf.usermodel.XSSFDrawing drawing,
      org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    for (org.apache.poi.xssf.usermodel.XSSFShape shape : drawing.getShapes()) {
      if (shapeUsesImagePart(shape, imagePartName)) {
        return true;
      }
    }
    return false;
  }

  private static boolean shapeUsesImagePart(
      org.apache.poi.xssf.usermodel.XSSFShape shape,
      org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    if (shape instanceof XSSFPicture picture) {
      return imagePartName.equals(
          ExcelDrawingPictureSupport.imagePartNameOrNull(picture).orElse(null));
    }
    if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
      Optional<org.apache.poi.openxml4j.opc.PackagePart> previewPart = previewImagePart(objectData);
      return previewPart.isPresent()
          && previewPart.orElseThrow().getPartName().equals(imagePartName);
    }
    return false;
  }

  public static void removeRelationshipsToPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart,
      org.apache.poi.openxml4j.opc.PackagePartName targetPartName) {
    if (sourcePart.isRelationshipPart()) {
      return;
    }
    List<String> relationshipIds = new ArrayList<>();
    try {
      for (org.apache.poi.openxml4j.opc.PackageRelationship relationship :
          sourcePart.getRelationships()) {
        if (relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
          continue;
        }
        if (targetPartName
            .getURI()
            .equals(
                org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
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

  public static void cleanupPackagePartIfUnused(
      org.apache.poi.openxml4j.opc.OPCPackage pkg,
      org.apache.poi.openxml4j.opc.PackagePartName partName) {
    ExcelPackageRelationshipSupport.cleanupPackagePartIfUnused(pkg, partName);
  }

  private static org.apache.poi.openxml4j.opc.PackagePart sheetPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return ((XSSFSheet) objectData.getDrawing().getParent()).getPackagePart();
  }

  private static org.apache.poi.poifs.filesystem.Ole10Native ole10Native(byte[] bytes)
      throws IOException, org.apache.poi.poifs.filesystem.Ole10NativeException {
    try (org.apache.poi.poifs.filesystem.POIFSFileSystem filesystem =
        new org.apache.poi.poifs.filesystem.POIFSFileSystem(new ByteArrayInputStream(bytes))) {
      org.apache.poi.poifs.filesystem.DirectoryNode directory = filesystem.getRoot();
      return org.apache.poi.poifs.filesystem.Ole10Native.createFromEmbeddedOleObject(directory);
    }
  }

  private record EmbeddedObjectReadback(
      ExcelEmbeddedObjectPackagingKind packagingKind,
      @Nullable String label,
      String fileName,
      @Nullable String command,
      String contentType,
      ExcelBinaryData payload,
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable ExcelBinaryData previewImage) {}
}
