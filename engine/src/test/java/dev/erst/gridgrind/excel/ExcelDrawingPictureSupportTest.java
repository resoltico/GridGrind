package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingPictureSupport;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Optional;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.w3c.dom.Node;

/** Focused regressions for malformed picture readback and recovery paths. */
class ExcelDrawingPictureSupportTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

  @Test
  void drawingObjectsAndPayloadFailClearlyWhenPictureRelationshipIsMissing() throws IOException {
    Path workbookPath = workbookWithPicture("gridgrind-picture-missing-relationship-");
    Path mutatedWorkbook =
        OoxmlPartMutator.rewriteEntry(
            workbookPath,
            "xl/drawings/drawing1.xml",
            xml -> xml.replaceFirst("r:embed=\"[^\"]+\"", "r:embed=\"rIdMissing\""));
    Path repairedWorkbook = XlsxRoundTrip.newWorkbookPath("gridgrind-picture-delete-recovery-");

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(mutatedWorkbook, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = workbook.sheet("Ops");

      IllegalStateException drawingObjectsFailure =
          assertThrows(IllegalStateException.class, () -> sheet.drawings().drawingObjects());
      assertTrue(drawingObjectsFailure.getMessage().contains("OpsPicture"));
      assertTrue(drawingObjectsFailure.getMessage().contains("rIdMissing"));

      IllegalStateException payloadFailure =
          assertThrows(
              IllegalStateException.class,
              () -> sheet.drawings().drawingObjectPayload("OpsPicture"));
      assertTrue(payloadFailure.getMessage().contains("OpsPicture"));
      assertTrue(payloadFailure.getMessage().contains("rIdMissing"));

      assertDoesNotThrow(() -> sheet.drawings().deleteDrawingObject("OpsPicture"));
      assertEquals(0, sheet.xssfSheet().getDrawingPatriarch().getShapes().size());

      workbook
          .persistence()
          .save(
              repairedWorkbook,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(repairedWorkbook, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(List.of(), reopened.sheet("Ops").drawings().drawingObjects());
    }
  }

  @Test
  void pictureHelpersHandleBlankMissingAndMismatchedRelationships() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      ExcelDrawingPictureSupport.PictureReadback readback =
          ExcelDrawingPictureSupport.requiredPictureReadback(picture);
      String relationId = ExcelDrawingPictureSupport.pictureRelationId(picture).orElseThrow();
      assertEquals(
          java.util.Optional.of(relationId),
          ExcelDrawingPictureSupport.reusableRelationId(
              picture.getDrawing(), relationId, readback.picturePart().getPartName()));
      assertTrue(
          ExcelDrawingPictureSupport.reusableRelationId(
                  picture.getDrawing(), null, readback.picturePart().getPartName())
              .isEmpty());
      assertTrue(
          ExcelDrawingPictureSupport.reusableRelationId(
                  picture.getDrawing(), " ", readback.picturePart().getPartName())
              .isEmpty());
      assertTrue(
          ExcelDrawingPictureSupport.reusableRelationId(
                  picture.getDrawing(),
                  relationId,
                  sheet.xssfSheet().getPackagePart().getPartName())
              .isEmpty());

      picture.getCTPicture().getBlipFill().getBlip().setEmbed("   ");
      assertTrue(ExcelDrawingPictureSupport.pictureRelationId(picture).isEmpty());
      assertTrue(ExcelDrawingPictureSupport.relatedImagePartOrNull(picture).isEmpty());
      assertTrue(ExcelDrawingPictureSupport.imagePartNameOrNull(picture).isEmpty());
      assertTrue(ExcelDrawingPictureSupport.pictureDataOrNull(picture).isEmpty());
      assertTrue(
          ExcelDrawingPictureSupport.missingPictureRelationship(picture)
              .getMessage()
              .contains("its image relationship"));
      IllegalArgumentException blankRelationId =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelDrawingPictureSupport.setPictureRelationId(picture, " "));
      assertEquals("relationId must not be blank", blankRelationId.getMessage());

      picture.getCTPicture().setBlipFill(CTBlipFillProperties.Factory.newInstance());
      IllegalStateException missingBlipPayload =
          assertThrows(
              IllegalStateException.class,
              () -> ExcelDrawingPictureSupport.setPictureRelationId(picture, "rId1"));
      assertTrue(missingBlipPayload.getMessage().contains("missing its blip payload"));

      XmlObject shapeXml = ExcelDrawingAnchorSupport.shapeXml(picture);
      ExcelDrawingController.LocatedShape located =
          new ExcelDrawingController.LocatedShape(
              picture.getDrawing(),
              picture,
              shapeXml,
              ExcelDrawingAnchorSupport.parentAnchor(shapeXml).orElse(null));
      assertDoesNotThrow(
          () -> ExcelDrawingPictureSupport.deletePicture(sheet.xssfSheet(), located, picture));
      assertEquals(0, sheet.xssfSheet().getDrawingPatriarch().getShapes().size());
    }
  }

  @Test
  void pictureDataFallbackUsesWorkbookCatalogWhenLoadedDrawingRelationDrifts() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      String detachedRelationId = "rIdDetached";
      String originalRelationId =
          ExcelDrawingPictureSupport.pictureRelationId(picture).orElseThrow();
      assertNotEquals(detachedRelationId, originalRelationId);

      ExcelDrawingPictureSupport.PictureReadback readback =
          ExcelDrawingPictureSupport.requiredPictureReadback(picture);
      picture
          .getDrawing()
          .getPackagePart()
          .addRelationship(
              readback.picturePart().getPartName(),
              TargetMode.INTERNAL,
              ExcelWorkbookImageCatalogSupport.pictureRelation(readback.format()).getRelation(),
              detachedRelationId);
      ExcelDrawingPictureSupport.setPictureRelationId(picture, detachedRelationId);

      assertNull(picture.getPictureData());
      assertTrue(ExcelDrawingPictureSupport.relatedImagePartOrNull(picture).isPresent());
      XSSFPictureData pictureData =
          ExcelDrawingPictureSupport.pictureDataOrNull(picture).orElseThrow();
      assertArrayEquals(PNG_PIXEL_BYTES, pictureData.getData());
      assertArrayEquals(
          PNG_PIXEL_BYTES, ExcelDrawingPictureSupport.requiredPictureReadback(picture).bytes());
    }
  }

  @Test
  void pictureDataFallbackRejectsNonImageRelationshipTargets() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      picture
          .getDrawing()
          .getPackagePart()
          .addRelationship(
              sheet.xssfSheet().getPackagePart().getPartName(),
              TargetMode.INTERNAL,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
              "rIdSheet");
      picture.getCTPicture().getBlipFill().getBlip().setEmbed("rIdSheet");

      assertNull(picture.getPictureData());
      assertEquals(
          sheet.xssfSheet().getPackagePart(),
          ExcelDrawingPictureSupport.relatedImagePartOrNull(picture).orElseThrow());
      assertTrue(ExcelDrawingPictureSupport.pictureDataOrNull(picture).isEmpty());
      assertTrue(ExcelDrawingPictureSupport.imagePartNameOrNull(picture).isEmpty());
      assertTrue(
          ExcelDrawingPictureSupport.missingPictureRelationship(picture)
              .getMessage()
              .contains("image relationship 'rIdSheet'"));
    }
  }

  @Test
  void pictureReadbackRejectsBlankContentType() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      ExcelDrawingPictureSupport.PictureReadback readback =
          ExcelDrawingPictureSupport.requiredPictureReadback(picture);

      IllegalArgumentException blankContentType =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  new ExcelDrawingPictureSupport.PictureReadback(
                      readback.pictureData(),
                      readback.picturePart(),
                      readback.format(),
                      " ",
                      readback.bytes()));
      assertEquals("contentType must not be blank", blankContentType.getMessage());
    }
  }

  @Test
  void pictureHelpersTreatMissingBlipFillAsMissingRelationship() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");

      removeBlipFill(picture);
      assertNull(picture.getCTPicture().getBlipFill());

      assertTrue(ExcelDrawingPictureSupport.pictureRelationId(picture).isEmpty());
      assertTrue(ExcelDrawingPictureSupport.pictureDataOrNull(picture).isEmpty());
      IllegalStateException missingBlipFill =
          assertThrows(
              IllegalStateException.class,
              () -> ExcelDrawingPictureSupport.setPictureRelationId(picture, "rId1"));
      assertTrue(missingBlipFill.getMessage().contains("missing its blip payload"));
    }
  }

  private static void removeBlipFill(XSSFPicture picture) {
    Node pictureNode = picture.getCTPicture().getDomNode();
    Node child = pictureNode.getFirstChild();
    while (child != null && !"blipFill".equals(child.getLocalName())) {
      child = child.getNextSibling();
    }
    assertNotNull(child);
    pictureNode.removeChild(child);
  }

  private static Path workbookWithPicture(String prefix) throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath(prefix);
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  Optional.of("Queue preview")));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    return workbookPath;
  }

  private static XSSFPicture requiredPicture(XSSFSheet sheet, String objectName) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFPicture.class::isInstance)
        .map(XSSFPicture.class::cast)
        .filter(shape -> objectName.equals(ExcelDrawingAnchorSupport.resolvedName(shape)))
        .findFirst()
        .orElseThrow();
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow, 0, 0),
        new ExcelDrawingMarker(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }
}
