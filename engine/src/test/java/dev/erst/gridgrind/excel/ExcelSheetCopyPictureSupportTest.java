package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;

/** Focused regressions for picture sheet-copy repair. */
class ExcelSheetCopyPictureSupportTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void repairCopiedPicturesRestoresCorruptedClonePictureRelationships() throws IOException {
    ExcelSheetCopyPictureSupport support = new ExcelSheetCopyPictureSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      sourceSheet.setPicture(pictureDefinition("OpsPicture2", 6, 1, 9, 6));

      ExcelSheetCopyPictureSupport.CopySnapshot snapshot =
          support.snapshot(sourceSheet.xssfSheet());
      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");

      XSSFSheet replicaPoiSheet = workbook.xssfWorkbook().getSheet("Replica");
      List<XSSFPicture> copiedPictures =
          replicaPoiSheet.createDrawingPatriarch().getShapes().stream()
              .filter(XSSFPicture.class::isInstance)
              .map(XSSFPicture.class::cast)
              .toList();
      assertEquals(2, copiedPictures.size());
      XSSFPicture corruptedPicture = requiredPicture(replicaPoiSheet, "OpsPicture2");
      ExcelDrawingPictureSupport.setPictureRelationId(corruptedPicture, "rIdMissing");
      assertNull(corruptedPicture.getPictureData());

      support.repairCopiedPictures(replicaPoiSheet, snapshot);

      XSSFPicture repairedPicture = requiredPicture(replicaPoiSheet, "OpsPicture2");
      assertTrue(copiedPictures.stream().allMatch(picture -> picture.getPictureData() != null));
      assertEquals(
          repairedPicture
              .getDrawing()
              .getRelationId(
                  ExcelDrawingPictureSupport.requiredPictureReadback(repairedPicture)
                      .pictureData()),
          repairedPicture.getCTPicture().getBlipFill().getBlip().getEmbed());
      ExcelDrawingObjectPayload.Picture payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.Picture.class,
              workbook.sheet("Replica").drawingObjectPayload("OpsPicture2"));
      assertArrayEquals(PNG_PIXEL_BYTES, payload.data().bytes());
    }
  }

  @Test
  void copySheetPreservesPicturesBeforeAndAfterRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-picture-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      sourceSheet.setPicture(pictureDefinition("OpsPicture2", 6, 1, 9, 6));

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      XSSFSheet replicaPoiSheet = workbook.xssfWorkbook().getSheet("Replica");
      assertTrue(
          replicaPoiSheet.getDrawingPatriarch().getShapes().stream()
              .filter(XSSFPicture.class::isInstance)
              .map(XSSFPicture.class::cast)
              .allMatch(picture -> picture.getPictureData() != null));
      assertEquals(
          List.of("OpsPicture1", "OpsPicture2"),
          workbook.sheet("Replica").drawingObjects().stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .filter(name -> name.startsWith("OpsPicture"))
              .toList());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals(
          List.of("OpsPicture1", "OpsPicture2"),
          reopened.sheet("Replica").drawingObjects().stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .filter(name -> name.startsWith("OpsPicture"))
              .toList());
      ExcelDrawingObjectPayload.Picture payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.Picture.class,
              reopened.sheet("Replica").drawingObjectPayload("OpsPicture2"));
      assertArrayEquals(PNG_PIXEL_BYTES, payload.data().bytes());
    }
  }

  @Test
  void repairCopiedPicturesRejectsMissingTargetDrawingPatriarch() throws IOException {
    ExcelSheetCopyPictureSupport support = new ExcelSheetCopyPictureSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      ExcelSheetCopyPictureSupport.CopySnapshot snapshot =
          support.snapshot(sourceSheet.xssfSheet());
      XSSFSheet targetSheet = workbook.getOrCreateSheet("Replica").xssfSheet();

      IllegalStateException missingDrawing =
          assertThrows(
              IllegalStateException.class,
              () -> support.repairCopiedPictures(targetSheet, snapshot));
      assertEquals(
          "Copied sheet 'Replica' is missing its drawing patriarch", missingDrawing.getMessage());
    }
  }

  @Test
  void repairCopiedPicturesRejectsPictureCountMismatch() throws IOException {
    ExcelSheetCopyPictureSupport support = new ExcelSheetCopyPictureSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      sourceSheet.setPicture(pictureDefinition("OpsPicture2", 6, 1, 9, 6));
      ExcelSheetCopyPictureSupport.CopySnapshot snapshot =
          support.snapshot(sourceSheet.xssfSheet());

      ExcelSheet targetSheet = workbook.getOrCreateSheet("Replica");
      targetSheet.setPicture(pictureDefinition("ReplicaPicture", 1, 1, 4, 6));

      IllegalStateException pictureCountMismatch =
          assertThrows(
              IllegalStateException.class,
              () -> support.repairCopiedPictures(targetSheet.xssfSheet(), snapshot));
      assertEquals(
          "Copied sheet 'Replica' changed its picture count during clone repair",
          pictureCountMismatch.getMessage());
    }
  }

  @Test
  void repairCopiedPicturesRejectsMissingAndInvalidSourceImageParts() throws Exception {
    ExcelSheetCopyPictureSupport support = new ExcelSheetCopyPictureSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet targetSheet = workbook.getOrCreateSheet("Replica");
      targetSheet.setPicture(pictureDefinition("ReplicaPicture", 1, 1, 4, 6));

      IllegalStateException missingImagePart =
          assertThrows(
              IllegalStateException.class,
              () ->
                  support.repairCopiedPictures(
                      targetSheet.xssfSheet(),
                      copySnapshotOf(
                          newPictureCopyPlan(
                              "ReplicaPicture", "/xl/media/missing.png", ExcelPictureFormat.PNG))));
      assertEquals(
          "Copied picture 'ReplicaPicture' is missing image part '/xl/media/missing.png'",
          missingImagePart.getMessage());

      IllegalStateException invalidPartName =
          assertThrows(
              IllegalStateException.class,
              () ->
                  support.repairCopiedPictures(
                      targetSheet.xssfSheet(),
                      copySnapshotOf(
                          newPictureCopyPlan(
                              "ReplicaPicture", "not-a-part-name", ExcelPictureFormat.PNG))));
      assertEquals(
          "Failed to resolve copied picture image part 'not-a-part-name'",
          invalidPartName.getMessage());
    }
  }

  @Test
  void pictureCopyPlanRejectsBlankRequiredFields() throws Exception {
    IllegalArgumentException blankObjectName =
        assertThrows(
            IllegalArgumentException.class,
            () -> newPictureCopyPlan(" ", "/xl/media/image1.png", ExcelPictureFormat.PNG));
    assertEquals("objectName must not be blank", blankObjectName.getMessage());

    IllegalArgumentException blankSourcePartName =
        assertThrows(
            IllegalArgumentException.class,
            () -> newPictureCopyPlan("ReplicaPicture", " ", ExcelPictureFormat.PNG));
    assertEquals("sourcePartName must not be blank", blankSourcePartName.getMessage());
  }

  private static ExcelPictureDefinition pictureDefinition(
      String pictureName, int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelPictureDefinition(
        pictureName,
        new ExcelBinaryData(PNG_PIXEL_BYTES),
        ExcelPictureFormat.PNG,
        anchor(fromColumn, fromRow, toColumn, toRow),
        "Queue preview");
  }

  private static XSSFPicture requiredPicture(XSSFSheet sheet, String objectName) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFPicture.class::isInstance)
        .map(XSSFPicture.class::cast)
        .filter(shape -> objectName.equals(ExcelDrawingAnchorSupport.resolvedName(shape)))
        .findFirst()
        .orElseThrow();
  }

  private static ExcelSheetCopyPictureSupport.CopySnapshot copySnapshotOf(
      ExcelSheetCopyPictureSupport.PictureCopyPlan... plans) {
    return new ExcelSheetCopyPictureSupport.CopySnapshot(List.of(plans));
  }

  private static ExcelSheetCopyPictureSupport.PictureCopyPlan newPictureCopyPlan(
      String objectName, String sourcePartName, ExcelPictureFormat format) {
    return new ExcelSheetCopyPictureSupport.PictureCopyPlan(objectName, sourcePartName, format);
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow, 0, 0),
        new ExcelDrawingMarker(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }
}
