package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.Provider;
import java.security.Security;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Drawing IO, digest, and cleanup failure coverage. */
class ExcelDrawingIoAndSecurityCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  void sha256FailureIsReportedWhenSecurityProvidersCannotSupplyTheDigest() throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();
    Provider[] providers = Security.getProviders();
    try {
      for (Provider provider : providers) {
        Security.removeProvider(provider.getName());
      }
      IllegalStateException failure =
          assertInvocationFailure(
              IllegalStateException.class,
              () -> invoke(controller, "sha256", String.class, PNG_PIXEL_BYTES));
      assertTrue(failure.getMessage().contains("SHA-256 digest is unavailable"));
    } finally {
      for (int index = 0; index < providers.length; index++) {
        Security.insertProviderAt(providers[index], index + 1);
      }
    }
  }

  @Test
  void ioSupportWrapsCheckedIoFailures() {
    assertEquals("ok", ExcelIoSupport.unchecked("boom", () -> "ok"));
    java.io.UncheckedIOException failure =
        assertThrows(
            java.io.UncheckedIOException.class,
            () ->
                ExcelIoSupport.unchecked(
                    "boom",
                    () -> {
                      throw new IOException("broken");
                    }));
    assertEquals("boom", failure.getMessage());
    assertEquals("broken", failure.getCause().getMessage());
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void drawingControllerCoversPackageCleanupAndEmbeddedFallbackBranches() throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFPicture picture = createPicture(workbook, drawing, "OpsPicture", 0, 0, 2, 2);
      XSSFObjectData objectData = createEmbeddedObject(workbook, drawing, "OpsEmbed", 3, 0, 6, 4);

      PackagePart previewPart =
          invoke(controller, "previewImagePart", PackagePart.class, objectData);
      assertTrue(
          invoke(
              controller,
              "imagePartUsed",
              Boolean.class,
              workbook,
              picture.getPictureData().getPackagePart().getPartName()));
      assertTrue(
          invoke(controller, "imagePartUsed", Boolean.class, workbook, previewPart.getPartName()));

      controller.setPicture(
          sheet,
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(0, 0, 2, 2),
              "replacement"));
      assertEquals(
          1L,
          controller.drawingObjects(sheet).stream()
              .filter(snapshot -> "OpsPicture".equals(snapshot.name()))
              .count());

      controller.setEmbeddedObject(
          sheet,
          new ExcelEmbeddedObjectDefinition(
              "OpsEmbed",
              "Payload",
              "payload.txt",
              "payload.txt",
              new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
              ExcelPictureFormat.PNG,
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              anchor(3, 0, 6, 4)));
      assertEquals(
          1L,
          controller.drawingObjects(sheet).stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData objectData =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "OpsEmbed", 0, 0, 3, 3);
      String relationId = objectData.getOleObject().getId();
      sheet.getPackagePart().removeRelationship(relationId);
      IllegalStateException missingPackageRelation =
          assertThrows(
              IllegalStateException.class,
              () -> controller.drawingObjectPayload(sheet, objectData.getShapeName()));
      assertTrue(missingPackageRelation.getMessage().contains("missing its package relationship"));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData objectData =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "OpsEmbed", 0, 0, 3, 3);
      String relationId = objectData.getOleObject().getId();
      PackageRelationship existingRelation = sheet.getPackagePart().getRelationship(relationId);
      sheet.getPackagePart().removeRelationship(relationId);
      ByteArrayOutputStream rawPayload = new ByteArrayOutputStream();
      rawPayload.write("raw".getBytes(StandardCharsets.UTF_8));
      PackagePart rawPart =
          workbook
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/embeddings/raw.bin"),
                  "application/octet-stream",
                  rawPayload);
      sheet
          .getPackagePart()
          .addRelationship(
              rawPart.getPartName(),
              TargetMode.INTERNAL,
              existingRelation.getRelationshipType(),
              relationId);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              controller.drawingObjectPayload(sheet, objectData.getShapeName()));
      assertEquals(ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE, payload.packagingKind());
      assertEquals("raw.bin", payload.fileName());
      assertArrayEquals("raw".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }

    try (OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream())) {
      PackagePart source =
          pkg.createPart(
              PackagingURIHelper.createPartName("/xl/workbook.xml"),
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
      pkg.createPart(
          PackagingURIHelper.createPartName("/xl/_rels/workbook.xml.rels"),
          "application/vnd.openxmlformats-package.relationships+xml");
      PackagePart usedPart =
          pkg.createPart(PackagingURIHelper.createPartName("/xl/media/used.png"), "image/png");
      PackagePart unusedPart =
          pkg.createPart(PackagingURIHelper.createPartName("/xl/media/unused.png"), "image/png");

      source.addRelationship(
          usedPart.getPartName(), TargetMode.INTERNAL, "urn:gridgrind:test", "rIdUsed");
      source.addRelationship(
          unusedPart.getPartName(), TargetMode.INTERNAL, "urn:gridgrind:test", "rIdUnused");
      source.addExternalRelationship(
          "https://example.com/object", "urn:gridgrind:test", "rIdExternal");

      invokeVoid(controller, "removeRelationshipsToPart", source, usedPart.getPartName());
      assertNull(source.getRelationship("rIdUsed"));
      assertNotNull(source.getRelationship("rIdUnused"));
      assertNotNull(source.getRelationship("rIdExternal"));

      invokeVoid(controller, "cleanupPackagePartIfUnused", pkg, usedPart.getPartName());
      assertFalse(pkg.containPart(usedPart.getPartName()));
      assertFalse(
          pkg.containPart(PackagingURIHelper.getRelationshipPartName(usedPart.getPartName())));
      invokeVoid(controller, "cleanupPackagePartIfUnused", pkg, unusedPart.getPartName());
      assertTrue(pkg.containPart(unusedPart.getPartName()));
      invokeVoid(controller, "cleanupPackagePartIfUnused", pkg, null);
      invokeVoid(
          controller,
          "cleanupPackagePartIfUnused",
          pkg,
          PackagingURIHelper.createPartName("/xl/media/missing.png"));
    }

    IllegalStateException invalidPackageFailure =
        assertInvocationFailure(
            IllegalStateException.class,
            () ->
                ExcelPackageRelationshipSupport.partIsStillReferenced(
                    List.of(
                        new InvalidRelationshipsPackagePart(
                            "/xl/worksheets/sheet1.xml",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")),
                    PackagingURIHelper.createPartName("/xl/media/unused.png")));
    assertTrue(
        invalidPackageFailure.getMessage().contains("Failed to inspect package relationships"));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData objectData =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "OpsEmbed", 0, 0, 3, 3);
      try (var cursor = objectData.getOleObject().newCursor()) {
        assertTrue(
            cursor.toChild(
                org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      objectData.getCTShape().getSpPr().unsetBlipFill();
      controller.deleteDrawingObject(sheet, objectData.getShapeName());
      assertEquals(0, controller.drawingObjects(sheet).size());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData objectData =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "OpsEmbed", 0, 0, 3, 3);
      String relationId = objectData.getOleObject().getId();
      PackageRelationship existingRelation = sheet.getPackagePart().getRelationship(relationId);
      sheet.getPackagePart().removeRelationship(relationId);
      ByteArrayOutputStream rawPayload = new ByteArrayOutputStream();
      rawPayload.write(ole2StorageBytes());
      PackagePart rawPart =
          workbook
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/embeddings/raw-ole.bin"),
                  "application/octet-stream",
                  rawPayload);
      sheet
          .getPackagePart()
          .addRelationship(
              rawPart.getPartName(),
              TargetMode.INTERNAL,
              existingRelation.getRelationshipType(),
              relationId);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              controller.drawingObjectPayload(sheet, objectData.getShapeName()));
      assertEquals(ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE, payload.packagingKind());
      assertEquals("raw-ole.bin", payload.fileName());
      assertArrayEquals(ole2StorageBytes(), payload.data().bytes());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData objectData =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "OpsEmbed", 0, 0, 3, 3);
      sheet.getPackagePart().removeRelationship(objectData.getOleObject().getId());
      controller.deleteDrawingObject(sheet, objectData.getShapeName());
      assertEquals(0, controller.drawingObjects(sheet).size());
    }
  }
}
