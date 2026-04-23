package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.List;
import java.util.Set;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
  void snapshotAndRepairHandleEmbeddedObjectsWithoutDrawingPreviewRelations() throws Exception {
    ExcelSheetCopyEmbeddedObjectSupport support = new ExcelSheetCopyEmbeddedObjectSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFObjectData sourceObject = requiredEmbeddedObject(sourceSheet.xssfSheet(), "OpsEmbed");
      sourceObject.getCTShape().getSpPr().unsetBlipFill();

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      assertEquals(1, snapshot.embeddedObjects().size());

      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");
      ExcelSheet replica = workbook.sheet("Replica");
      support.repairCopiedEmbeddedObjects(replica, snapshot);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              replica.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void repairCopiedEmbeddedObjectsSkipsDrawingPreviewRepairWhenCopiedObjectLosesPreviewBlip()
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
      if (copiedObject.getCTShape().getSpPr().isSetBlipFill()) {
        copiedObject.getCTShape().getSpPr().unsetBlipFill();
      }
      assertNull(ExcelDrawingBinarySupport.previewDrawingRelationId(copiedObject));

      support.repairCopiedEmbeddedObjects(replica, snapshot);

      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              replica.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void repairCopiedEmbeddedObjectsSkipsDrawingPreviewRepairWhenSnapshotLacksDrawingPreviewPart()
      throws Exception {
    ExcelSheetCopyEmbeddedObjectSupport support = new ExcelSheetCopyEmbeddedObjectSupport();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFObjectData sourceObject = requiredEmbeddedObject(sourceSheet.xssfSheet(), "OpsEmbed");
      String previewDrawingRelationId =
          ExcelDrawingBinarySupport.previewDrawingRelationId(sourceObject);
      assertNotNull(previewDrawingRelationId);
      sourceObject.getCTShape().getSpPr().unsetBlipFill();

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      sourceObject
          .getCTShape()
          .getSpPr()
          .addNewBlipFill()
          .addNewBlip()
          .setEmbed(previewDrawingRelationId);

      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");
      ExcelSheet replica = workbook.sheet("Replica");
      XSSFObjectData copiedObject = requiredEmbeddedObject(replica.xssfSheet(), "OpsEmbed");
      assertNotNull(ExcelDrawingBinarySupport.previewDrawingRelationId(copiedObject));

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

  @Test
  void nextWorksheetRelationIdRespectsReservedIdsAndWrapsRelationshipInspectionFailures()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      var worksheet = sheet.xssfSheet().getCTWorksheet();
      worksheet.addNewLegacyDrawingHF().setId("rIdHeaderLegacy");
      worksheet.addNewDrawingHF().setId("rIdHeaderFooter");

      String nextId =
          ExcelSheetCopyEmbeddedObjectSupport.nextWorksheetRelationId(
              sheet.xssfSheet(), () -> List.of());
      assertNotEquals("rIdHeaderLegacy", nextId);
      assertNotEquals("rIdHeaderFooter", nextId);

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.nextWorksheetRelationId(
                      sheet.xssfSheet(),
                      () -> {
                        throw new InvalidFormatException("broken relationships");
                      }));
      assertTrue(failure.getMessage().contains("Failed to inspect worksheet relationships"));
    }
  }

  @Test
  void worksheetRelationHelpersRecognizeWorksheetAndOtherEmbeddedObjectReferences()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      XSSFSheet blankSheet = workbook.xssfWorkbook().createSheet("Blank");
      blankSheet.getCTWorksheet().addNewDrawing().setId("rIdDrawing");
      blankSheet.getCTWorksheet().addNewLegacyDrawing().setId("rIdLegacy");
      assertTrue(
          ExcelSheetCopyEmbeddedObjectSupport.worksheetStructureReferencesId(
              blankSheet.getCTWorksheet(), "rIdDrawing"));
      assertTrue(
          ExcelSheetCopyEmbeddedObjectSupport.worksheetStructureReferencesId(
              blankSheet.getCTWorksheet(), "rIdLegacy"));
      Set<String> blankReferencedIds =
          ExcelSheetCopyEmbeddedObjectSupport.referencedWorksheetRelationIds(
              workbook.xssfWorkbook().createSheet("ReallyBlank"));
      assertEquals(Set.of(), blankReferencedIds);

      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsA", "alpha"));
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsB", "beta"));

      XSSFSheet poiSheet = sheet.xssfSheet();
      XSSFObjectData firstObject = requiredEmbeddedObject(poiSheet, "OpsA");
      XSSFObjectData secondObject = requiredEmbeddedObject(poiSheet, "OpsB");
      assertFalse(
          worksheetRelationIdReferencedElsewhere(
              workbook.xssfWorkbook().createSheet("NoOleObjects"),
              firstObject,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              "rIdMissing"));
      var worksheet = poiSheet.getCTWorksheet();
      worksheet.addNewLegacyDrawingHF().setId("rIdHeaderLegacy");
      worksheet.addNewDrawingHF().setId("rIdHeaderFooter");

      assertTrue(
          ExcelSheetCopyEmbeddedObjectSupport.worksheetStructureReferencesId(
              worksheet, "rIdHeaderLegacy"));
      assertTrue(
          ExcelSheetCopyEmbeddedObjectSupport.worksheetStructureReferencesId(
              worksheet, "rIdHeaderFooter"));

      Set<String> referencedIds =
          ExcelSheetCopyEmbeddedObjectSupport.referencedWorksheetRelationIds(poiSheet);
      assertTrue(referencedIds.contains("rIdHeaderLegacy"));
      assertTrue(referencedIds.contains("rIdHeaderFooter"));

      assertTrue(
          worksheetRelationIdReferencedElsewhere(
              poiSheet,
              firstObject,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              ExcelDrawingBinarySupport.nullIfBlank(secondObject.getOleObject().getId())));
      assertTrue(
          worksheetRelationIdReferencedElsewhere(
              poiSheet,
              firstObject,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.PREVIEW_SHEET,
              ExcelDrawingBinarySupport.previewSheetRelationId(secondObject.getOleObject())));
    }
  }

  @Test
  void repairSheetDrawingRelationCoversEarlyReturnsAndMissingPatriarch() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet noDrawing = workbook.createSheet("NoDrawing");
      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(noDrawing);

      XSSFSheet blankDrawingId = workbook.createSheet("BlankDrawingId");
      blankDrawingId.getCTWorksheet().addNewDrawing().setId(" ");
      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(blankDrawingId);

      XSSFSheet missingPatriarch = workbook.createSheet("Ghost");
      missingPatriarch.getCTWorksheet().addNewDrawing().setId("rIdGhost");
      IllegalStateException missingDrawing =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(missingPatriarch));
      assertTrue(missingDrawing.getMessage().contains("missing its drawing patriarch"));
    }
  }

  @Test
  void repairSheetDrawingRelationRebindsDrawingRelations()
      throws IOException, InvalidFormatException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFSheet poiSheet = sheet.xssfSheet();
      var drawingPatriarch = poiSheet.createDrawingPatriarch();
      String originalDrawingRelationId = poiSheet.getCTWorksheet().getDrawing().getId();
      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(poiSheet, drawingPatriarch);

      poiSheet.getCTWorksheet().getDrawing().setId(" ");
      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(poiSheet, drawingPatriarch);
      poiSheet.getCTWorksheet().getDrawing().setId(originalDrawingRelationId);

      XSSFObjectData objectData = requiredEmbeddedObject(poiSheet, "OpsEmbed");
      String conflictingId = objectData.getOleObject().getId();
      poiSheet.getCTWorksheet().getDrawing().setId(conflictingId);

      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(poiSheet, drawingPatriarch);

      assertEquals(
          XSSFRelation.DRAWINGS.getRelation(),
          poiSheet.getPackagePart().getRelationship(conflictingId).getRelationshipType());
      assertEquals(
          drawingPatriarch.getPackagePart().getPartName(),
          ExcelDrawingBinarySupport.relatedInternalPart(poiSheet.getPackagePart(), conflictingId)
              .getPartName());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet source = workbook.getOrCreateSheet("Source");
      source.setEmbeddedObject(embeddedObjectDefinition("SourceEmbed", "payload"));
      ExcelSheet target = workbook.getOrCreateSheet("Target");
      target.setEmbeddedObject(embeddedObjectDefinition("TargetEmbed", "payload"));

      XSSFSheet sourceSheet = source.xssfSheet();
      XSSFSheet targetSheet = target.xssfSheet();
      var sourceDrawing = sourceSheet.createDrawingPatriarch();
      var targetDrawing = targetSheet.createDrawingPatriarch();

      String targetDrawingRelationId = targetSheet.getCTWorksheet().getDrawing().getId();
      targetSheet
          .getPackagePart()
          .addRelationship(
              sourceDrawing.getPackagePart().getPartName(),
              TargetMode.INTERNAL,
              XSSFRelation.DRAWINGS.getRelation(),
              "rIdForeignDrawing");
      targetSheet.getCTWorksheet().getDrawing().setId("rIdForeignDrawing");

      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(targetSheet, targetDrawing);
      assertEquals(
          targetDrawing.getPackagePart().getPartName(),
          ExcelDrawingBinarySupport.relatedInternalPart(
                  targetSheet.getPackagePart(), targetSheet.getCTWorksheet().getDrawing().getId())
              .getPartName());

      targetSheet.getCTWorksheet().getDrawing().setId(targetDrawingRelationId + "Missing");
      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(targetSheet, targetDrawing);
      assertEquals(
          targetDrawing.getPackagePart().getPartName(),
          ExcelDrawingBinarySupport.relatedInternalPart(
                  targetSheet.getPackagePart(), targetSheet.getCTWorksheet().getDrawing().getId())
              .getPartName());

      targetSheet
          .getPackagePart()
          .addRelationship(
              PackagingURIHelper.createPartName("/xl/drawings/missing-drawing.xml"),
              TargetMode.INTERNAL,
              XSSFRelation.DRAWINGS.getRelation(),
              "rIdMissingPart");
      targetSheet.getCTWorksheet().getDrawing().setId("rIdMissingPart");
      ExcelSheetCopyEmbeddedObjectSupport.repairSheetDrawingRelation(targetSheet, targetDrawing);
      assertEquals(
          targetDrawing.getPackagePart().getPartName(),
          ExcelDrawingBinarySupport.relatedInternalPart(
                  targetSheet.getPackagePart(), targetSheet.getCTWorksheet().getDrawing().getId())
              .getPartName());
    }
  }

  @Test
  void repairWorksheetBoundRelationReusesIdsForReplaceableAndMissingRelations()
      throws IOException, InvalidFormatException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFSheet poiSheet = sheet.xssfSheet();
      XSSFObjectData objectData = requiredEmbeddedObject(poiSheet, "OpsEmbed");
      String relationId = objectData.getOleObject().getId();
      ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot sourcePart =
          ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
              poiSheet.getPackagePart(), relationId, "OpsEmbed", "embedded object package");

      PackagePart existingPart =
          ExcelDrawingBinarySupport.relatedInternalPart(poiSheet.getPackagePart(), relationId);
      PackagePartName existingPartName = existingPart.getPartName();
      try (var outputStream = existingPart.getOutputStream()) {
        outputStream.write("wrong".getBytes(StandardCharsets.UTF_8));
      }

      String repairedId =
          ExcelSheetCopyEmbeddedObjectSupport.repairWorksheetBoundRelation(
              poiSheet,
              objectData,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              relationId,
              sourcePart,
              "embedded object package",
              "OpsEmbed");
      assertEquals(relationId, repairedId);
      PackagePart repairedPart =
          ExcelDrawingBinarySupport.relatedInternalPart(poiSheet.getPackagePart(), repairedId);
      assertNotNull(repairedPart);
      assertNotEquals(existingPartName, repairedPart.getPartName());
      assertArrayEquals(sourcePart.bytes().bytes(), repairedPart.getInputStream().readAllBytes());
      assertFalse(workbook.xssfWorkbook().getPackage().containPart(existingPartName));

      poiSheet.getPackagePart().removeRelationship(repairedId);
      String restoredMissingId =
          ExcelSheetCopyEmbeddedObjectSupport.repairWorksheetBoundRelation(
              poiSheet,
              objectData,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              repairedId,
              sourcePart,
              "embedded object package",
              "OpsEmbed");
      assertEquals(repairedId, restoredMissingId);
      assertNotNull(
          ExcelDrawingBinarySupport.relatedInternalPart(
              poiSheet.getPackagePart(), restoredMissingId));
    }
  }

  @Test
  void repairWorksheetBoundRelationReusesIdsWhenRelationshipsPointAtMissingParts()
      throws IOException, InvalidFormatException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFSheet poiSheet = sheet.xssfSheet();
      XSSFObjectData objectData = requiredEmbeddedObject(poiSheet, "OpsEmbed");
      ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot sourcePart =
          ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
              poiSheet.getPackagePart(),
              objectData.getOleObject().getId(),
              "OpsEmbed",
              "embedded object package");

      poiSheet
          .getPackagePart()
          .addRelationship(
              PackagingURIHelper.createPartName("/xl/embeddings/missing-part.bin"),
              TargetMode.INTERNAL,
              sourcePart.relationshipType(),
              "rIdMissingPart");
      objectData.getOleObject().setId("rIdMissingPart");

      String repairedId =
          ExcelSheetCopyEmbeddedObjectSupport.repairWorksheetBoundRelation(
              poiSheet,
              objectData,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              "rIdMissingPart",
              sourcePart,
              "embedded object package",
              "OpsEmbed");
      assertEquals("rIdMissingPart", repairedId);
      assertNotNull(
          ExcelDrawingBinarySupport.relatedInternalPart(poiSheet.getPackagePart(), repairedId));
    }
  }

  @Test
  void repairWorksheetBoundRelationReallocatesWrongRelationTargets()
      throws IOException, InvalidFormatException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      XSSFSheet poiSheet = sheet.xssfSheet();
      XSSFObjectData objectData = requiredEmbeddedObject(poiSheet, "OpsEmbed");
      ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot sourcePart =
          ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
              poiSheet.getPackagePart(),
              objectData.getOleObject().getId(),
              "OpsEmbed",
              "embedded object package");

      PackagePart bogusImage =
          workbook
              .xssfWorkbook()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/media/gridgrind-bogus.png"), "image/png");
      try (var outputStream = bogusImage.getOutputStream()) {
        outputStream.write(PNG_PIXEL_BYTES);
      }
      poiSheet
          .getPackagePart()
          .addRelationship(
              bogusImage.getPartName(),
              TargetMode.INTERNAL,
              XSSFRelation.IMAGES.getRelation(),
              "rIdBogus");
      objectData.getOleObject().setId("rIdBogus");

      String imageRepairedId =
          ExcelSheetCopyEmbeddedObjectSupport.repairWorksheetBoundRelation(
              poiSheet,
              objectData,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              "rIdBogus",
              sourcePart,
              "embedded object package",
              "OpsEmbed");
      assertNotEquals("rIdBogus", imageRepairedId);
      assertEquals(
          sourcePart.relationshipType(),
          poiSheet.getPackagePart().getRelationship(imageRepairedId).getRelationshipType());
      assertArrayEquals(
          sourcePart.bytes().bytes(),
          ExcelDrawingBinarySupport.relatedInternalPart(poiSheet.getPackagePart(), imageRepairedId)
              .getInputStream()
              .readAllBytes());

      PackagePart wrongContentTypePart =
          workbook
              .xssfWorkbook()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/embeddings/gridgrind-wrong.bin"),
                  "application/octet-stream");
      try (var outputStream = wrongContentTypePart.getOutputStream()) {
        outputStream.write(sourcePart.bytes().bytes());
      }
      poiSheet
          .getPackagePart()
          .addRelationship(
              wrongContentTypePart.getPartName(),
              TargetMode.INTERNAL,
              sourcePart.relationshipType(),
              "rIdWrongContent");
      objectData.getOleObject().setId("rIdWrongContent");

      String contentTypeRepairedId =
          ExcelSheetCopyEmbeddedObjectSupport.repairWorksheetBoundRelation(
              poiSheet,
              objectData,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              "rIdWrongContent",
              sourcePart,
              "embedded object package",
              "OpsEmbed");
      assertNotEquals("rIdWrongContent", contentTypeRepairedId);
      assertArrayEquals(
          sourcePart.bytes().bytes(),
          ExcelDrawingBinarySupport.relatedInternalPart(
                  poiSheet.getPackagePart(), contentTypeRepairedId)
              .getInputStream()
              .readAllBytes());
    }
  }

  @Test
  void repairWorksheetBoundRelationReallocatesIdsOwnedByAnotherObject()
      throws IOException, InvalidFormatException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsA", "alpha"));
      sheet.setEmbeddedObject(embeddedObjectDefinition("OpsB", "beta"));

      XSSFSheet poiSheet = sheet.xssfSheet();
      XSSFObjectData firstObject = requiredEmbeddedObject(poiSheet, "OpsA");
      XSSFObjectData secondObject = requiredEmbeddedObject(poiSheet, "OpsB");
      String secondRelationId = secondObject.getOleObject().getId();
      ExcelSheetCopyEmbeddedObjectSupport.InternalRelationSnapshot secondSourcePart =
          ExcelSheetCopyEmbeddedObjectSupport.requiredInternalRelation(
              poiSheet.getPackagePart(), secondRelationId, "OpsB", "embedded object package");

      String repairedId =
          ExcelSheetCopyEmbeddedObjectSupport.repairWorksheetBoundRelation(
              poiSheet,
              firstObject,
              ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole.OLE_OBJECT,
              secondRelationId,
              secondSourcePart,
              "embedded object package",
              "OpsA");
      assertNotEquals(secondRelationId, repairedId);
      assertArrayEquals(
          secondSourcePart.bytes().bytes(),
          ExcelDrawingBinarySupport.relatedInternalPart(poiSheet.getPackagePart(), repairedId)
              .getInputStream()
              .readAllBytes());
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

  private static boolean worksheetRelationIdReferencedElsewhere(
      XSSFSheet sheet,
      XSSFObjectData objectData,
      ExcelSheetCopyEmbeddedObjectSupport.WorksheetRelationRole relationRole,
      String relationId) {
    return ExcelSheetCopyEmbeddedObjectSupport.worksheetRelationIdReferencedElsewhere(
        sheet, objectData, relationRole, relationId);
  }
}
