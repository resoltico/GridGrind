package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.List;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.junit.jupiter.api.Test;

/** Branch coverage for embedded-object sheet-copy repair helpers. */
class ExcelSheetCopyEmbeddedObjectSupportCoverageTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void snapshotHandlesBlankSheetsPicturesAndPreviewlessEmbeddedObjects() throws Exception {
    ExcelSheetCopyEmbeddedObjectSupport support = new ExcelSheetCopyEmbeddedObjectSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet blankSheet = workbook.getOrCreateSheet("Blank");
      assertEquals(List.of(), support.snapshot(blankSheet).embeddedObjects());
      support.repairCopiedEmbeddedObjects(
          blankSheet, new ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot(List.of()));

      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      createPicture(workbook.xssfWorkbook(), sourceSheet.xssfSheet(), "OpsPicture");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFObjectData sourceObject = requiredEmbeddedObject(sourceSheet.xssfSheet(), "OpsEmbed");
      try (XmlCursor cursor = sourceObject.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      assertEquals(1, snapshot.embeddedObjects().size());

      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");
      ExcelSheet replica = workbook.sheet("Replica");

      support.repairCopiedEmbeddedObjects(replica, snapshot);
      support.repairCopiedEmbeddedObjects(replica, snapshot);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              replica.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void repairCopiedEmbeddedObjectsSkipsPreviewRepairWhenCopiedObjectOmitsObjectPr()
      throws Exception {
    ExcelSheetCopyEmbeddedObjectSupport support = new ExcelSheetCopyEmbeddedObjectSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");
      ExcelSheet replica = workbook.sheet("Replica");

      XSSFObjectData copiedObject = requiredEmbeddedObject(replica.xssfSheet(), "OpsEmbed");
      try (XmlCursor cursor = copiedObject.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }

      support.repairCopiedEmbeddedObjects(replica, snapshot);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              replica.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void repairCopiedEmbeddedObjectsSkipsPreviewRepairWhenSnapshotLacksPreviewPart()
      throws Exception {
    ExcelSheetCopyEmbeddedObjectSupport support = new ExcelSheetCopyEmbeddedObjectSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFObjectData sourceObject = requiredEmbeddedObject(sourceSheet.xssfSheet(), "OpsEmbed");
      try (XmlCursor cursor = sourceObject.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");
      ExcelSheet replica = workbook.sheet("Replica");

      XSSFObjectData copiedObject = requiredEmbeddedObject(replica.xssfSheet(), "OpsEmbed");
      addObjectPreviewReference(copiedObject, "rIdSyntheticPreview");

      support.repairCopiedEmbeddedObjects(replica, snapshot);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              replica.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void packageVisibleHelperFailuresSurfaceClearEmbeddedObjectDiagnostics() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet blankSheet = workbook.createSheet("Blank");
      IllegalStateException missingDrawing =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredEmbeddedObject(
                      blankSheet, "OpsEmbed"));
      assertTrue(missingDrawing.getMessage().contains("missing its drawing patriarch"));

      XSSFSheet pictureOnlySheet = workbook.createSheet("PictureOnly");
      createPicture(workbook, pictureOnlySheet, "OpsPicture");
      IllegalStateException missingObject =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredEmbeddedObject(
                      pictureOnlySheet, "OpsEmbed"));
      assertTrue(missingObject.getMessage().contains("missing embedded object 'OpsEmbed'"));

      PackagePart sheetPart = blankSheet.getPackagePart();
      IllegalStateException nullId =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
                      sheetPart, null, "OpsEmbed", "embedded object package"));
      assertTrue(nullId.getMessage().contains("missing its embedded object package id"));

      IllegalStateException missingId =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
                      sheetPart, " ", "OpsEmbed", "embedded object package"));
      assertTrue(missingId.getMessage().contains("missing its embedded object package id"));

      IllegalStateException missingRelationship =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
                      sheetPart, "rIdMissing", "OpsEmbed", "embedded object package"));
      assertTrue(
          missingRelationship
              .getMessage()
              .contains("missing its embedded object package relationship"));

      sheetPart.addExternalRelationship(
          "https://example.com/object", "urn:gridgrind:test", "rIdExt");
      IllegalStateException externalRelationship =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
                      sheetPart, "rIdExt", "OpsEmbed", "embedded object package"));
      assertTrue(
          externalRelationship
              .getMessage()
              .contains("missing its embedded object package relationship"));

      sheetPart.addRelationship(
          PackagingURIHelper.createPartName("/xl/embeddings/missing.bin"),
          TargetMode.INTERNAL,
          "urn:gridgrind:test",
          "rIdBroken");
      IllegalStateException missingPart =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
                      sheetPart, "rIdBroken", "OpsEmbed", "embedded object package"));
      assertTrue(missingPart.getMessage().contains("missing its embedded object package part"));
    }
  }

  @Test
  void packageVisiblePartCopyHelpersCoverFreshAndInvalidTargetPartPaths() throws Exception {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot(
                " ", "application/octet-stream", "/xl/embeddings/item.bin", binary("payload")));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot relationSnapshot =
          new ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot(
              "urn:gridgrind:test",
              "application/octet-stream",
              "/xl/embeddings/item.bin",
              binary("payload"));

      ExcelSheetCopyEmbeddedObjectSupport.ensureInternalRelation(
          sheet.getPackagePart(),
          "rIdPayload",
          relationSnapshot,
          "embedded object package",
          "OpsEmbed");
      PackagePart copiedPart =
          ExcelDrawingBinarySupport.relatedInternalPart(sheet.getPackagePart(), "rIdPayload");
      assertNotNull(copiedPart);
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8), copiedPart.getInputStream().readAllBytes());

      ExcelSheetCopyEmbeddedObjectSupport.ensureInternalRelation(
          sheet.getPackagePart(),
          "rIdPayload",
          relationSnapshot,
          "embedded object package",
          "OpsEmbed");

      sheet
          .getPackagePart()
          .addRelationship(
              PackagingURIHelper.createPartName("/xl/embeddings/missing.bin"),
              TargetMode.INTERNAL,
              "urn:gridgrind:test",
              "rIdBrokenPayload");
      ExcelSheetCopyEmbeddedObjectSupport.ensureInternalRelation(
          sheet.getPackagePart(),
          "rIdBrokenPayload",
          relationSnapshot,
          "embedded object package",
          "OpsEmbed");
      PackagePart repairedBrokenPart =
          ExcelDrawingBinarySupport.relatedInternalPart(sheet.getPackagePart(), "rIdBrokenPayload");
      assertNotNull(repairedBrokenPart);
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          repairedBrokenPart.getInputStream().readAllBytes());
    }

    try (OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream())) {
      pkg.createPart(
          PackagingURIHelper.createPartName("/xl/embeddings/item-gridgrind-copy-1.bin"),
          "application/octet-stream");
      PackagePartName nextPartName =
          ExcelSheetCopyEmbeddedObjectSupport.nextCopiedPartName(pkg, "/xl/embeddings/item.bin");
      assertEquals("/xl/embeddings/item-gridgrind-copy-2.bin", nextPartName.getName());

      ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot invalidSnapshot =
          new ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot(
              "urn:gridgrind:test", "application/octet-stream", "invalid", binary("payload"));
      IllegalStateException invalidPartName =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.createCopiedPart(
                      pkg, invalidSnapshot, "embedded object package", "OpsEmbed"));
      assertTrue(invalidPartName.getMessage().contains("Failed to copy embedded object package"));
    }
  }

  private static void createPicture(XSSFWorkbook workbook, XSSFSheet sheet, String objectName) {
    int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
    var drawing = sheet.createDrawingPatriarch();
    drawing
        .createPicture(drawing.createAnchor(0, 0, 0, 0, 0, 0, 2, 2), pictureIndex)
        .getCTPicture()
        .getNvPicPr()
        .getCNvPr()
        .setName(objectName);
  }

  private static ExcelEmbeddedObjectDefinition embeddedObjectDefinition(
      String objectName, String payloadText) {
    return new ExcelEmbeddedObjectDefinition(
        objectName,
        "Payload",
        "payload.txt",
        "payload.txt",
        binary(payloadText),
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(PNG_PIXEL_BYTES),
        new ExcelDrawingAnchor.TwoCell(
            new ExcelDrawingMarker(1, 1, 0, 0),
            new ExcelDrawingMarker(4, 6, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));
  }

  private static ExcelBinaryData binary(String text) {
    return new ExcelBinaryData(text.getBytes(StandardCharsets.UTF_8));
  }

  private static XSSFObjectData requiredEmbeddedObject(XSSFSheet sheet, String objectName) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFObjectData.class::isInstance)
        .map(XSSFObjectData.class::cast)
        .filter(shape -> objectName.equals(ExcelDrawingAnchorSupport.resolvedName(shape)))
        .findFirst()
        .orElseThrow();
  }

  private static void addObjectPreviewReference(XSSFObjectData objectData, String relationId) {
    try (XmlCursor cursor = objectData.getOleObject().newCursor()) {
      cursor.toEndToken();
      cursor.beginElement("objectPr", XSSFRelation.NS_SPREADSHEETML);
      cursor.insertAttributeWithValue(
          new QName(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS, "id"), relationId);
    }
  }
}
