package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Base64;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;

/** Focused regressions for embedded-object sheet-copy repair. */
class ExcelSheetCopyEmbeddedObjectSupportTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void repairCopiedEmbeddedObjectsRestoresPoiCloneSheetPackageRelationships() throws IOException {
    ExcelSheetCopyEmbeddedObjectSupport support = new ExcelSheetCopyEmbeddedObjectSupport();
    ExcelDrawingController drawingController = new ExcelDrawingController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");

      XSSFSheet replicaPoiSheet = workbook.xssfWorkbook().getSheet("Replica");
      XSSFObjectData copiedObject = requiredEmbeddedObject(replicaPoiSheet, "OpsEmbed");
      assertNull(drawingController.oleObjectPart(copiedObject));

      support.repairCopiedEmbeddedObjects(workbook.sheet("Replica"), snapshot);

      XSSFObjectData repairedObject = requiredEmbeddedObject(replicaPoiSheet, "OpsEmbed");
      assertNotNull(drawingController.oleObjectPart(repairedObject));
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Replica").drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void copySheetPreservesEmbeddedObjectPayloadsBeforeAndAfterRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-embedded-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      ExcelDrawingObjectPayload.EmbeddedObject copiedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Replica").drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), copiedPayload.data().bytes());
      assertEquals(
          1L,
          workbook.sheet("Replica").drawingObjects().stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      ExcelDrawingObjectPayload.EmbeddedObject copiedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              reopened.sheet("Replica").drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), copiedPayload.data().bytes());
      assertEquals(
          1L,
          reopened.sheet("Replica").drawingObjects().stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());
    }
  }

  private static ExcelEmbeddedObjectDefinition embeddedObjectDefinition(
      String objectName, String payloadText) {
    return new ExcelEmbeddedObjectDefinition(
        objectName,
        "Payload",
        "payload.txt",
        "payload.txt",
        new ExcelBinaryData(payloadText.getBytes(StandardCharsets.UTF_8)),
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(PNG_PIXEL_BYTES),
        anchor(1, 1, 4, 6));
  }

  private static XSSFObjectData requiredEmbeddedObject(XSSFSheet sheet, String objectName) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFObjectData.class::isInstance)
        .map(XSSFObjectData.class::cast)
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
