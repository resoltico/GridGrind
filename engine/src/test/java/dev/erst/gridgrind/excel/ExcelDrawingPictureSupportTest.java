package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Map;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
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
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void drawingObjectsAndPayloadFailClearlyWhenPictureRelationshipIsMissing() throws IOException {
    Path workbookPath = workbookWithPicture("gridgrind-picture-missing-relationship-");
    Path mutatedWorkbook =
        OoxmlPartMutator.rewriteEntry(
            workbookPath,
            "xl/drawings/drawing1.xml",
            xml -> xml.replaceFirst("r:embed=\"[^\"]+\"", "r:embed=\"rIdMissing\""));
    Path repairedWorkbook = XlsxRoundTrip.newWorkbookPath("gridgrind-picture-delete-recovery-");

    try (ExcelWorkbook workbook = ExcelWorkbook.open(mutatedWorkbook)) {
      ExcelSheet sheet = workbook.sheet("Ops");

      IllegalStateException drawingObjectsFailure =
          assertThrows(IllegalStateException.class, sheet::drawingObjects);
      assertTrue(drawingObjectsFailure.getMessage().contains("OpsPicture"));
      assertTrue(drawingObjectsFailure.getMessage().contains("rIdMissing"));

      IllegalStateException payloadFailure =
          assertThrows(IllegalStateException.class, () -> sheet.drawingObjectPayload("OpsPicture"));
      assertTrue(payloadFailure.getMessage().contains("OpsPicture"));
      assertTrue(payloadFailure.getMessage().contains("rIdMissing"));

      assertDoesNotThrow(() -> sheet.deleteDrawingObject("OpsPicture"));
      assertEquals(0, sheet.xssfSheet().getDrawingPatriarch().getShapes().size());

      workbook.save(repairedWorkbook);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(repairedWorkbook)) {
      assertEquals(List.of(), reopened.sheet("Ops").drawingObjects());
    }
  }

  @Test
  void pictureHelpersHandleBlankMissingAndMismatchedRelationships() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      ExcelDrawingPictureSupport.PictureReadback readback =
          ExcelDrawingPictureSupport.requiredPictureReadback(picture);
      String relationId = ExcelDrawingPictureSupport.pictureRelationId(picture);
      assertNotNull(relationId);
      assertEquals(
          relationId,
          ExcelDrawingPictureSupport.reusableRelationId(
              picture.getDrawing(), relationId, readback.picturePart().getPartName()));
      assertNull(
          ExcelDrawingPictureSupport.reusableRelationId(
              picture.getDrawing(), null, readback.picturePart().getPartName()));
      assertNull(
          ExcelDrawingPictureSupport.reusableRelationId(
              picture.getDrawing(), " ", readback.picturePart().getPartName()));
      assertNull(
          ExcelDrawingPictureSupport.reusableRelationId(
              picture.getDrawing(), relationId, sheet.xssfSheet().getPackagePart().getPartName()));

      picture.getCTPicture().getBlipFill().getBlip().setEmbed("   ");
      assertNull(ExcelDrawingPictureSupport.pictureRelationId(picture));
      assertNull(ExcelDrawingPictureSupport.relatedImagePartOrNull(picture));
      assertNull(ExcelDrawingPictureSupport.imagePartNameOrNull(picture));
      assertNull(ExcelDrawingPictureSupport.pictureDataOrNull(picture));
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
              ExcelDrawingAnchorSupport.parentAnchor(shapeXml));
      assertDoesNotThrow(
          () -> ExcelDrawingPictureSupport.deletePicture(sheet.xssfSheet(), located, picture));
      assertEquals(0, sheet.xssfSheet().getDrawingPatriarch().getShapes().size());
    }
  }

  @Test
  void pictureDataFallbackUsesWorkbookCatalogWhenLoadedDrawingRelationDrifts() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      String relationId = ExcelDrawingPictureSupport.pictureRelationId(picture);
      assertNotNull(relationId);

      removeLoadedRelation(picture.getDrawing(), relationId);

      assertNull(picture.getPictureData());
      assertNotNull(ExcelDrawingPictureSupport.relatedImagePartOrNull(picture));
      XSSFPictureData pictureData = ExcelDrawingPictureSupport.pictureDataOrNull(picture);
      assertNotNull(pictureData);
      assertArrayEquals(PNG_PIXEL_BYTES, pictureData.getData());
      assertArrayEquals(
          PNG_PIXEL_BYTES, ExcelDrawingPictureSupport.requiredPictureReadback(picture).bytes());
    }
  }

  @Test
  void pictureDataFallbackRejectsNonImageRelationshipTargets() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
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
          ExcelDrawingPictureSupport.relatedImagePartOrNull(picture));
      assertNull(ExcelDrawingPictureSupport.pictureDataOrNull(picture));
      assertNull(ExcelDrawingPictureSupport.imagePartNameOrNull(picture));
      assertTrue(
          ExcelDrawingPictureSupport.missingPictureRelationship(picture)
              .getMessage()
              .contains("image relationship 'rIdSheet'"));
    }
  }

  @Test
  void pictureReadbackRejectsBlankContentType() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");
      ExcelDrawingPictureSupport.PictureReadback readback =
          ExcelDrawingPictureSupport.requiredPictureReadback(picture);

      IllegalArgumentException blankContentType =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  newPictureReadback(
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
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      XSSFPicture picture = requiredPicture(sheet.xssfSheet(), "OpsPicture");

      removeBlipFill(picture);
      assertNull(picture.getCTPicture().getBlipFill());

      assertNull(ExcelDrawingPictureSupport.pictureRelationId(picture));
      assertNull(ExcelDrawingPictureSupport.pictureDataOrNull(picture));
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
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      workbook.save(workbookPath);
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

  @SuppressWarnings({"unchecked", "PMD.AvoidAccessibilityAlteration"})
  private static void removeLoadedRelation(XSSFDrawing drawing, String relationId) {
    try {
      Field relationsField = POIXMLDocumentPart.class.getDeclaredField("relations");
      relationsField.setAccessible(true);
      Map<String, ?> relations = (Map<String, ?>) relationsField.get(drawing);
      relations.remove(relationId);
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
  }

  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  private static ExcelDrawingPictureSupport.PictureReadback newPictureReadback(
      XSSFPictureData pictureData,
      org.apache.poi.openxml4j.opc.PackagePart picturePart,
      ExcelPictureFormat format,
      String contentType,
      byte[] bytes) {
    try {
      Constructor<ExcelDrawingPictureSupport.PictureReadback> constructor =
          ExcelDrawingPictureSupport.PictureReadback.class.getDeclaredConstructor(
              XSSFPictureData.class,
              org.apache.poi.openxml4j.opc.PackagePart.class,
              ExcelPictureFormat.class,
              String.class,
              byte[].class);
      constructor.setAccessible(true);
      return constructor.newInstance(pictureData, picturePart, format, contentType, bytes);
    } catch (InvocationTargetException exception) {
      Throwable cause = exception.getCause();
      if (cause instanceof RuntimeException runtime) {
        runtime.addSuppressed(exception);
        throw runtime;
      }
      AssertionError assertion = new AssertionError(cause);
      assertion.addSuppressed(exception);
      throw assertion;
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow, 0, 0),
        new ExcelDrawingMarker(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }
}
