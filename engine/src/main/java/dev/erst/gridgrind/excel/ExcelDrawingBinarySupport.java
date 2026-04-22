package dev.erst.gridgrind.excel;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Objects;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Binary drawing payload, preview-image, and embedded-object package helpers. */
final class ExcelDrawingBinarySupport {
  private static final javax.xml.namespace.QName OLE_PREVIEW_ID =
      new javax.xml.namespace.QName(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS, "id");

  private ExcelDrawingBinarySupport() {}

  static ExcelDrawingObjectSnapshot.EmbeddedObject snapshotEmbeddedObject(
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

  static ExcelDrawingObjectPayload.Picture picturePayload(String objectName, XSSFPicture picture) {
    org.apache.poi.xssf.usermodel.XSSFPictureData pictureData = picture.getPictureData();
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

  static ExcelDrawingObjectPayload.EmbeddedObject embeddedObjectPayload(
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

  static org.apache.poi.openxml4j.opc.PackagePart previewImagePart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    String sheetRelationId = previewSheetRelationId(objectData.getOleObject());
    if (sheetRelationId != null) {
      org.apache.poi.openxml4j.opc.PackagePart previewPart =
          relatedInternalPart(sheetPart(objectData), sheetRelationId);
      if (previewPart != null) {
        return previewPart;
      }
    }
    String drawingRelationId = previewDrawingRelationId(objectData);
    return drawingRelationId == null
        ? null
        : relatedInternalPart(objectData.getDrawing().getPackagePart(), drawingRelationId);
  }

  static org.apache.poi.openxml4j.opc.PackagePart relatedInternalPart(
      org.apache.poi.openxml4j.opc.PackagePart sourcePart, String relationshipId) {
    if (relationshipId == null || relationshipId.isBlank()) {
      return null;
    }
    org.apache.poi.openxml4j.opc.PackageRelationship relationship =
        sourcePart.getRelationship(relationshipId);
    if (relationship == null
        || relationship.getTargetMode() == org.apache.poi.openxml4j.opc.TargetMode.EXTERNAL) {
      return null;
    }
    try {
      return sourcePart
          .getPackage()
          .getPart(
              org.apache.poi.openxml4j.opc.PackagingURIHelper.createPartName(
                  org.apache.poi.openxml4j.opc.PackagingURIHelper.resolvePartUri(
                      sourcePart.getPartName().getURI(), relationship.getTargetURI())));
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to resolve related package part", exception);
    }
  }

  static void cleanupWorkbookImagePartIfUnused(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    if (imagePartName == null || imagePartUsed(workbook, imagePartName)) {
      return;
    }
    removeRelationshipsToPart(workbook.getPackagePart(), imagePartName);
    cleanupPackagePartIfUnused(workbook.getPackage(), imagePartName);
  }

  static void removeOleObject(
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

  @SuppressWarnings("PMD.UseTryWithResources")
  static String previewSheetRelationId(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject oleObject) {
    org.apache.xmlbeans.XmlCursor cursor = oleObject.newCursor();
    try {
      return cursor.toChild(org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML, "objectPr")
          ? cursor.getAttributeText(OLE_PREVIEW_ID)
          : null;
    } finally {
      cursor.close();
    }
  }

  static String previewDrawingRelationId(org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    if (objectData.getCTShape().getSpPr().getBlipFill() == null
        || objectData.getCTShape().getSpPr().getBlipFill().getBlip() == null) {
      return null;
    }
    return objectData.getCTShape().getSpPr().getBlipFill().getBlip().getEmbed();
  }

  static ExcelBinaryData binary(byte[] bytes, String label) {
    try {
      return new ExcelBinaryData(bytes);
    } catch (IllegalArgumentException exception) {
      throw new IllegalStateException("Missing " + label + " bytes", exception);
    }
  }

  static byte[] partBytes(org.apache.poi.openxml4j.opc.PackagePart part) {
    try {
      return part.getInputStream().readAllBytes();
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read package part bytes", exception);
    }
  }

  static String sha256(byte[] bytes) {
    try {
      MessageDigest digest = MessageDigest.getInstance("SHA-256");
      return HexFormat.of().formatHex(digest.digest(bytes));
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 digest is unavailable", exception);
    }
  }

  static String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }

  static String firstNonBlank(String first, String second) {
    String normalizedFirst = nullIfBlank(first);
    return normalizedFirst != null ? normalizedFirst : nullIfBlank(second);
  }

  static boolean looksLikeOle2Storage(byte[] bytes) throws IOException {
    try (ByteArrayInputStream input = new ByteArrayInputStream(bytes)) {
      return org.apache.poi.poifs.filesystem.FileMagic.valueOf(
              org.apache.poi.poifs.filesystem.FileMagic.prepareToCheckMagic(input))
          == org.apache.poi.poifs.filesystem.FileMagic.OLE2;
    }
  }

  static String partFileName(org.apache.poi.openxml4j.opc.PackagePart part) {
    String partName = part.getPartName().getName();
    return partName.substring(partName.lastIndexOf('/') + 1);
  }

  private static EmbeddedObjectReadback embeddedObjectReadback(
      String objectName, org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    org.apache.poi.openxml4j.opc.PackagePart previewPart = previewImagePart(objectData);
    ExcelBinaryData previewImage =
        previewPart == null ? null : binary(partBytes(previewPart), "preview image");
    ExcelPictureFormat previewFormat =
        previewPart == null
            ? null
            : ExcelPictureFormat.fromContentType(previewPart.getContentType());
    org.apache.poi.openxml4j.opc.PackagePart objectPart = oleObjectPart(objectData);
    if (objectPart == null) {
      throw new IllegalStateException(
          "Embedded object '" + objectName + "' is missing its package relationship");
    }
    ExcelBinaryData rawPackage = binary(partBytes(objectPart), "embedded object package");
    String contentType = objectPart.getContentType();
    String label = null;
    String fileName = partFileName(objectPart);
    String command = null;
    ExcelEmbeddedObjectPackagingKind packagingKind = ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE;
    ExcelBinaryData payload = rawPackage;
    try {
      if (looksLikeOle2Storage(rawPackage.bytes())) {
        org.apache.poi.poifs.filesystem.Ole10Native nativeData = ole10Native(rawPackage.bytes());
        payload = new ExcelBinaryData(nativeData.getDataBuffer());
        label = firstNonBlank(nativeData.getLabel2(), nativeData.getLabel());
        fileName = firstNonBlank(nativeData.getFileName2(), nativeData.getFileName());
        command = firstNonBlank(nativeData.getCommand2(), nativeData.getCommand());
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

  private static org.apache.poi.openxml4j.opc.PackagePart oleObjectPart(
      org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
    return relatedInternalPart(sheetPart(objectData), objectData.getOleObject().getId());
  }

  static boolean imagePartUsed(
      XSSFWorkbook workbook, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    ExcelSignatureLineController signatureLineController = new ExcelSignatureLineController();
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
      org.apache.poi.xssf.usermodel.XSSFDrawing drawing = sheet.getDrawingPatriarch();
      if (drawing == null) {
        if (signatureLineController.usesImagePart(sheet, imagePartName)) {
          return true;
        }
      } else {
        for (org.apache.poi.xssf.usermodel.XSSFShape shape : drawing.getShapes()) {
          if (shape instanceof XSSFPicture picture
              && picture.getPictureData().getPackagePart().getPartName().equals(imagePartName)) {
            return true;
          }
          if (shape instanceof org.apache.poi.xssf.usermodel.XSSFObjectData objectData) {
            org.apache.poi.openxml4j.opc.PackagePart previewPart = previewImagePart(objectData);
            if (previewPart != null && previewPart.getPartName().equals(imagePartName)) {
              return true;
            }
          }
        }
        if (signatureLineController.usesImagePart(sheet, imagePartName)) {
          return true;
        }
      }
    }
    return false;
  }

  static void removeRelationshipsToPart(
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

  static void cleanupPackagePartIfUnused(
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
      String label,
      String fileName,
      String command,
      String contentType,
      ExcelBinaryData payload,
      ExcelPictureFormat previewFormat,
      ExcelBinaryData previewImage) {}
}
