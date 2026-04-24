package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Picture-specific relation, readback, and deletion helpers. */
final class ExcelDrawingPictureSupport {
  private ExcelDrawingPictureSupport() {}

  static ExcelDrawingObjectSnapshot.Picture snapshotPicture(XSSFPicture picture) {
    PictureReadback readback = requiredPictureReadback(picture);
    ExcelDrawingSnapshotSupport.RasterDimensions dimensions =
        ExcelDrawingSnapshotSupport.rasterDimensions(readback.bytes());
    return new ExcelDrawingObjectSnapshot.Picture(
        ExcelDrawingAnchorSupport.resolvedName(picture),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(picture)),
        readback.format(),
        readback.contentType(),
        readback.bytes().length,
        ExcelDrawingBinarySupport.sha256(readback.bytes()),
        dimensions.widthPixels(),
        dimensions.heightPixels(),
        ExcelDrawingBinarySupport.nullIfBlank(
            picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()));
  }

  static ExcelDrawingObjectPayload.Picture picturePayload(String objectName, XSSFPicture picture) {
    PictureReadback readback = requiredPictureReadback(picture);
    return new ExcelDrawingObjectPayload.Picture(
        objectName,
        readback.format(),
        readback.contentType(),
        objectName + readback.format().defaultExtension(),
        ExcelDrawingBinarySupport.sha256(readback.bytes()),
        ExcelBinaryData.readback(readback.bytes()),
        ExcelDrawingBinarySupport.nullIfBlank(
            picture.getCTPicture().getNvPicPr().getCNvPr().getDescr()));
  }

  static void deletePicture(
      XSSFSheet sheet, ExcelDrawingController.LocatedShape located, XSSFPicture picture) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(located, "located must not be null");
    Objects.requireNonNull(picture, "picture must not be null");

    PackagePartName imagePartName = imagePartNameOrNull(picture).orElse(null);
    Optional<String> relationId = pictureRelationId(picture);
    if (relationId.isPresent()) {
      located.drawing().getPackagePart().removeRelationship(relationId.orElseThrow());
    }
    ExcelDrawingRemovalSupport.removeParentAnchor(located.drawing(), located.parentAnchor());
    if (imagePartName != null) {
      ExcelDrawingBinarySupport.cleanupWorkbookImagePartIfUnused(
          sheet.getWorkbook(), imagePartName);
    }
  }

  static PictureReadback requiredPictureReadback(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    XSSFPictureData pictureData = pictureDataOrNull(picture).orElse(null);
    if (pictureData == null) {
      throw missingPictureRelationship(picture);
    }
    byte[] bytes = pictureData.getData();
    return new PictureReadback(
        pictureData,
        pictureData.getPackagePart(),
        ExcelPicturePoiBridge.fromPoiPictureType(pictureData.getPictureType()),
        pictureData.getPackagePart().getContentType(),
        bytes);
  }

  static Optional<XSSFPictureData> pictureDataOrNull(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    String relationId = pictureRelationId(picture).orElse(null);
    if (relationId == null) {
      return Optional.empty();
    }
    XSSFPictureData pictureData = picture.getPictureData();
    if (pictureData != null) {
      return Optional.of(pictureData);
    }
    PackagePart imagePart =
        ExcelDrawingBinarySupport.relatedInternalPart(
            picture.getDrawing().getPackagePart(), relationId);
    if (imagePart == null) {
      return Optional.empty();
    }
    try {
      ExcelPictureFormat.fromContentType(imagePart.getContentType());
    } catch (IllegalArgumentException ignored) {
      return Optional.empty();
    }
    return Optional.of(ExcelWorkbookImageCatalogSupport.pictureDataForPart(imagePart));
  }

  static Optional<PackagePart> relatedImagePartOrNull(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    String relationId = pictureRelationId(picture).orElse(null);
    if (relationId == null) {
      return Optional.empty();
    }
    XSSFPictureData pictureData = picture.getPictureData();
    if (pictureData != null) {
      return Optional.of(pictureData.getPackagePart());
    }
    return Optional.ofNullable(
        ExcelDrawingBinarySupport.relatedInternalPart(
            picture.getDrawing().getPackagePart(), relationId));
  }

  static Optional<PackagePartName> imagePartNameOrNull(XSSFPicture picture) {
    PackagePart imagePart = relatedImagePartOrNull(picture).orElse(null);
    if (imagePart == null) {
      return Optional.empty();
    }
    try {
      ExcelPictureFormat.fromContentType(imagePart.getContentType());
      return Optional.of(imagePart.getPartName());
    } catch (IllegalArgumentException ignored) {
      return Optional.empty();
    }
  }

  static Optional<String> pictureRelationId(XSSFPicture picture) {
    Objects.requireNonNull(picture, "picture must not be null");
    if (picture.getCTPicture().getBlipFill() == null
        || picture.getCTPicture().getBlipFill().getBlip() == null) {
      return Optional.empty();
    }
    return Optional.ofNullable(
        ExcelDrawingBinarySupport.nullIfBlank(
            picture.getCTPicture().getBlipFill().getBlip().getEmbed()));
  }

  static void setPictureRelationId(XSSFPicture picture, String relationId) {
    Objects.requireNonNull(picture, "picture must not be null");
    Objects.requireNonNull(relationId, "relationId must not be null");
    if (relationId.isBlank()) {
      throw new IllegalArgumentException("relationId must not be blank");
    }
    if (picture.getCTPicture().getBlipFill() == null
        || picture.getCTPicture().getBlipFill().getBlip() == null) {
      throw new IllegalStateException(
          "Picture '"
              + ExcelDrawingAnchorSupport.resolvedName(picture)
              + "' on sheet '"
              + sheetName(picture)
              + "' is missing its blip payload");
    }
    picture.getCTPicture().getBlipFill().getBlip().setEmbed(relationId);
  }

  static Optional<String> reusableRelationId(
      org.apache.poi.xssf.usermodel.XSSFDrawing drawing,
      String relationId,
      PackagePartName targetPartName) {
    Objects.requireNonNull(drawing, "drawing must not be null");
    Objects.requireNonNull(targetPartName, "targetPartName must not be null");
    if (relationId == null || relationId.isBlank()) {
      return Optional.empty();
    }
    RelationPart existingRelation = drawing.getRelationPartById(relationId);
    if (existingRelation == null) {
      return Optional.of(relationId);
    }
    return existingRelation.getDocumentPart().getPackagePart().getPartName().equals(targetPartName)
        ? Optional.of(relationId)
        : Optional.empty();
  }

  static IllegalStateException missingPictureRelationship(XSSFPicture picture) {
    String relationId = pictureRelationId(picture).orElse(null);
    String relationFragment =
        relationId == null ? "its image relationship" : "image relationship '" + relationId + "'";
    return new IllegalStateException(
        "Picture '"
            + ExcelDrawingAnchorSupport.resolvedName(picture)
            + "' on sheet '"
            + sheetName(picture)
            + "' is missing "
            + relationFragment);
  }

  private static String sheetName(XSSFPicture picture) {
    return ((XSSFSheet) picture.getDrawing().getParent()).getSheetName();
  }

  /** Resolved picture binary plus enough metadata to rebuild factual readback snapshots. */
  static final class PictureReadback {
    private final XSSFPictureData pictureData;
    private final PackagePart picturePart;
    private final ExcelPictureFormat format;
    private final String contentType;
    private final byte[] bytes;

    PictureReadback(
        XSSFPictureData pictureData,
        PackagePart picturePart,
        ExcelPictureFormat format,
        String contentType,
        byte[] bytes) {
      Objects.requireNonNull(pictureData, "pictureData must not be null");
      Objects.requireNonNull(picturePart, "picturePart must not be null");
      Objects.requireNonNull(format, "format must not be null");
      Objects.requireNonNull(contentType, "contentType must not be null");
      if (contentType.isBlank()) {
        throw new IllegalArgumentException("contentType must not be blank");
      }
      this.pictureData = pictureData;
      this.picturePart = picturePart;
      this.format = format;
      this.contentType = contentType;
      this.bytes = Objects.requireNonNull(bytes, "bytes must not be null").clone();
    }

    XSSFPictureData pictureData() {
      return pictureData;
    }

    PackagePart picturePart() {
      return picturePart;
    }

    ExcelPictureFormat format() {
      return format;
    }

    String contentType() {
      return contentType;
    }

    byte[] bytes() {
      return bytes.clone();
    }
  }
}
