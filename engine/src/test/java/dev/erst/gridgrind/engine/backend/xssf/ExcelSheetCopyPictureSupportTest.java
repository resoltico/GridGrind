package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;

/** Focused regressions for picture sheet-copy repair. */
class ExcelSheetCopyPictureSupportTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

  @Test
  void repairCopiedPicturesRestoresCorruptedClonePictureRelationships() throws IOException {
    SheetCopyAdapter support = new SheetCopyAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture2", 6, 1, 9, 6));

      SheetCopyAdapter.PictureCopySnapshot snapshot = support.snapshot(sourceSheet.xssfSheet());
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
      DrawingAdapter.setPictureRelationId(corruptedPicture, "rIdMissing");
      assertNull(corruptedPicture.getPictureData());

      support.repairCopiedPictures(replicaPoiSheet, snapshot);

      XSSFPicture repairedPicture = requiredPicture(replicaPoiSheet, "OpsPicture2");
      assertTrue(copiedPictures.stream().allMatch(picture -> picture.getPictureData() != null));
      assertEquals(
          repairedPicture
              .getDrawing()
              .getRelationId(DrawingAdapter.requiredPictureReadback(repairedPicture).pictureData()),
          repairedPicture.getCTPicture().getBlipFill().getBlip().getEmbed());
      DrawingObjectPayloadReport.Picture payload =
          assertInstanceOf(
              DrawingObjectPayloadReport.Picture.class,
              workbook.sheet("Replica").drawings().drawingObjectPayload("OpsPicture2"));
      assertArrayEquals(
          PNG_PIXEL_BYTES, java.util.Base64.getDecoder().decode(payload.base64Data()));
    }
  }

  @Test
  void copySheetPreservesPicturesBeforeAndAfterRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-picture-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture2", 6, 1, 9, 6));

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      XSSFSheet replicaPoiSheet = workbook.xssfWorkbook().getSheet("Replica");
      assertTrue(
          replicaPoiSheet.getDrawingPatriarch().getShapes().stream()
              .filter(XSSFPicture.class::isInstance)
              .map(XSSFPicture.class::cast)
              .allMatch(picture -> picture.getPictureData() != null));
      assertEquals(
          List.of("OpsPicture1", "OpsPicture2"),
          workbook.sheet("Replica").drawings().drawingObjects().stream()
              .map(DrawingObjectReport::name)
              .filter(name -> name.startsWith("OpsPicture"))
              .toList());

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          List.of("OpsPicture1", "OpsPicture2"),
          reopened.sheet("Replica").drawings().drawingObjects().stream()
              .map(DrawingObjectReport::name)
              .filter(name -> name.startsWith("OpsPicture"))
              .toList());
      DrawingObjectPayloadReport.Picture payload =
          assertInstanceOf(
              DrawingObjectPayloadReport.Picture.class,
              reopened.sheet("Replica").drawings().drawingObjectPayload("OpsPicture2"));
      assertArrayEquals(
          PNG_PIXEL_BYTES, java.util.Base64.getDecoder().decode(payload.base64Data()));
    }
  }

  @Test
  void repairCopiedPicturesRejectsMissingTargetDrawingPatriarch() throws IOException {
    SheetCopyAdapter support = new SheetCopyAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      SheetCopyAdapter.PictureCopySnapshot snapshot = support.snapshot(sourceSheet.xssfSheet());
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
    SheetCopyAdapter support = new SheetCopyAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture1", 1, 1, 4, 6));
      sourceSheet.drawings().setPicture(pictureDefinition("OpsPicture2", 6, 1, 9, 6));
      SheetCopyAdapter.PictureCopySnapshot snapshot = support.snapshot(sourceSheet.xssfSheet());

      XssfSheet targetSheet = workbook.getOrCreateSheet("Replica");
      targetSheet.drawings().setPicture(pictureDefinition("ReplicaPicture", 1, 1, 4, 6));

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
    SheetCopyAdapter support = new SheetCopyAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet targetSheet = workbook.getOrCreateSheet("Replica");
      targetSheet.drawings().setPicture(pictureDefinition("ReplicaPicture", 1, 1, 4, 6));

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

  private static ResolvedPicture pictureDefinition(
      String pictureName, int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ResolvedPicture(
        pictureName,
        new BinaryContent(PNG_PIXEL_BYTES),
        ExcelPictureFormat.PNG,
        anchor(fromColumn, fromRow, toColumn, toRow),
        Optional.of("Queue preview"));
  }

  private static XSSFPicture requiredPicture(XSSFSheet sheet, String objectName) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFPicture.class::isInstance)
        .map(XSSFPicture.class::cast)
        .filter(shape -> objectName.equals(DrawingAdapter.resolvedName(shape)))
        .findFirst()
        .orElseThrow();
  }

  private static SheetCopyAdapter.PictureCopySnapshot copySnapshotOf(
      SheetCopyAdapter.PictureCopyPlan... plans) {
    return new SheetCopyAdapter.PictureCopySnapshot(List.of(plans));
  }

  private static SheetCopyAdapter.PictureCopyPlan newPictureCopyPlan(
      String objectName, String sourcePartName, ExcelPictureFormat format) {
    return new SheetCopyAdapter.PictureCopyPlan(objectName, sourcePartName, format);
  }

  private static DrawingAnchorSpec.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromColumn, fromRow, 0, 0),
        new DrawingMarkerSpec(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }
}
