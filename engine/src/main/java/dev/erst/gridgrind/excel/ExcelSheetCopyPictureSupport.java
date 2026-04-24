package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ooxml.POIXMLDocumentPart.RelationPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Repairs copied picture relation ids after POI sheet cloning leaves stale source drawing ids in
 * the cloned picture XML.
 */
final class ExcelSheetCopyPictureSupport {
  CopySnapshot snapshot(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    XSSFDrawing drawing = sourceSheet.getDrawingPatriarch();
    if (drawing == null) {
      return new CopySnapshot(List.of());
    }

    List<PictureCopyPlan> pictures = new ArrayList<>();
    for (var shape : drawing.getShapes()) {
      if (shape instanceof XSSFPicture picture) {
        pictures.add(snapshotPicture(picture));
      }
    }
    return new CopySnapshot(pictures);
  }

  void repairCopiedPictures(XSSFSheet targetSheet, CopySnapshot snapshot) {
    Objects.requireNonNull(targetSheet, "targetSheet must not be null");
    Objects.requireNonNull(snapshot, "snapshot must not be null");
    if (snapshot.pictures().isEmpty()) {
      return;
    }

    XSSFDrawing drawing = targetSheet.getDrawingPatriarch();
    if (drawing == null) {
      throw new IllegalStateException(
          "Copied sheet '" + targetSheet.getSheetName() + "' is missing its drawing patriarch");
    }
    List<XSSFPicture> targetPictures =
        drawing.getShapes().stream()
            .filter(XSSFPicture.class::isInstance)
            .map(XSSFPicture.class::cast)
            .toList();
    if (targetPictures.size() != snapshot.pictures().size()) {
      throw new IllegalStateException(
          "Copied sheet '"
              + targetSheet.getSheetName()
              + "' changed its picture count during clone repair");
    }

    for (int index = 0; index < targetPictures.size(); index++) {
      repairCopiedPicture(targetPictures.get(index), snapshot.pictures().get(index));
    }
  }

  private static PictureCopyPlan snapshotPicture(XSSFPicture picture) {
    String objectName = ExcelDrawingAnchorSupport.resolvedName(picture);
    ExcelDrawingPictureSupport.PictureReadback readback =
        ExcelDrawingPictureSupport.requiredPictureReadback(picture);
    return new PictureCopyPlan(
        objectName, readback.picturePart().getPartName().getName(), readback.format());
  }

  private static void repairCopiedPicture(
      XSSFPicture targetPicture, PictureCopyPlan sourcePicture) {
    PackagePart imagePart =
        requiredImagePart(targetPicture.getDrawing().getPackagePart(), sourcePicture);
    String preferredRelationId =
        ExcelDrawingPictureSupport.reusableRelationId(
                targetPicture.getDrawing(),
                ExcelDrawingPictureSupport.pictureRelationId(targetPicture).orElse(null),
                imagePart.getPartName())
            .orElse(null);
    RelationPart relation =
        targetPicture
            .getDrawing()
            .addRelation(
                preferredRelationId,
                ExcelWorkbookImageCatalogSupport.pictureRelation(sourcePicture.format()),
                ExcelWorkbookImageCatalogSupport.pictureDataForPart(imagePart));
    ExcelDrawingPictureSupport.setPictureRelationId(
        targetPicture, relation.getRelationship().getId());
  }

  private static PackagePart requiredImagePart(
      org.apache.poi.openxml4j.opc.PackagePart drawingPart, PictureCopyPlan sourcePicture) {
    try {
      PackagePartName partName = PackagingURIHelper.createPartName(sourcePicture.sourcePartName());
      PackagePart imagePart = drawingPart.getPackage().getPart(partName);
      if (imagePart == null) {
        throw new IllegalStateException(
            "Copied picture '"
                + sourcePicture.objectName()
                + "' is missing image part '"
                + sourcePicture.sourcePartName()
                + "'");
      }
      return imagePart;
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(
          "Failed to resolve copied picture image part '" + sourcePicture.sourcePartName() + "'",
          exception);
    }
  }

  record CopySnapshot(List<PictureCopyPlan> pictures) {
    CopySnapshot {
      pictures = List.copyOf(pictures);
    }
  }

  record PictureCopyPlan(
      String objectName,
      String sourcePartName,
      dev.erst.gridgrind.excel.foundation.ExcelPictureFormat format) {
    PictureCopyPlan {
      Objects.requireNonNull(objectName, "objectName must not be null");
      if (objectName.isBlank()) {
        throw new IllegalArgumentException("objectName must not be blank");
      }
      Objects.requireNonNull(sourcePartName, "sourcePartName must not be null");
      if (sourcePartName.isBlank()) {
        throw new IllegalArgumentException("sourcePartName must not be blank");
      }
      Objects.requireNonNull(format, "format must not be null");
    }
  }
}
