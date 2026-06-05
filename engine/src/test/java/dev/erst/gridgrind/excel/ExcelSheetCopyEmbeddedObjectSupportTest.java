package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingBinarySupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Base64;
import java.util.Optional;
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

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      ExcelSheetCopyEmbeddedObjectSupport.CopySnapshot snapshot = support.snapshot(sourceSheet);
      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");

      XSSFSheet replicaPoiSheet = workbook.xssfWorkbook().getSheet("Replica");
      XSSFObjectData copiedObject = requiredEmbeddedObject(replicaPoiSheet, "OpsEmbed");
      assertEquals(Optional.empty(), ExcelDrawingBinarySupport.oleObjectPart(copiedObject));
      assertNull(previewSheetPart(copiedObject));
      assertNull(previewDrawingPart(copiedObject));

      support.repairCopiedEmbeddedObjects(workbook.sheet("Replica"), snapshot);

      XSSFObjectData repairedObject = requiredEmbeddedObject(replicaPoiSheet, "OpsEmbed");
      assertTrue(ExcelDrawingBinarySupport.oleObjectPart(repairedObject).isPresent());
      assertNotNull(previewSheetPart(repairedObject));
      assertNotNull(previewDrawingPart(repairedObject));
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Replica").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void copySheetPreservesEmbeddedObjectPayloadsBeforeAndAfterRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-embedded-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      workbook.sheets().copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());

      ExcelDrawingObjectPayload.EmbeddedObject copiedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Replica").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), copiedPayload.data().bytes());
      assertEquals(
          1L,
          workbook.sheet("Replica").drawings().drawingObjects().stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());

      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectPayload.EmbeddedObject copiedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              reopened.sheet("Replica").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), copiedPayload.data().bytes());
      assertEquals(
          1L,
          reopened.sheet("Replica").drawings().drawingObjects().stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());
    }
  }

  @Test
  void copySheetPreservesEmbeddedObjectsWhenCommentsAndRepeatedSheetCreatesShiftWorksheetIds()
      throws IOException {
    Path workbookPath =
        XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-embedded-comment-relations-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sourceSheet = workbook.getOrCreateSheet("LL");
      sourceSheet
          .annotations()
          .setComment("C11", new ExcelComment("Note Name3", "GridGrind", false));
      sourceSheet.drawings().setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      workbook.sheets().copySheet("LL", "LL_B1", new ExcelSheetCopyPosition.AtIndex(0));
      workbook.getOrCreateSheet("LL");
      workbook.getOrCreateSheet("LL");
      workbook.getOrCreateSheet("LL");
      workbook.getOrCreateSheet("LL");

      ExcelDrawingObjectPayload.EmbeddedObject copiedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("LL_B1").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), copiedPayload.data().bytes());
      workbook.persistence().save(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelDrawingObjectPayload.EmbeddedObject copiedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              reopened.sheet("LL_B1").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), copiedPayload.data().bytes());
      assertEquals(
          1L,
          reopened.sheet("LL_B1").drawings().drawingObjects().stream()
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

  private static org.apache.poi.openxml4j.opc.PackagePart previewDrawingPart(
      XSSFObjectData objectData) {
    return ExcelDrawingBinarySupport.previewDrawingRelationId(objectData)
        .flatMap(
            relationId ->
                ExcelDrawingBinarySupport.relatedInternalPart(
                    objectData.getDrawing().getPackagePart(), relationId))
        .orElse(null);
  }

  private static org.apache.poi.openxml4j.opc.PackagePart previewSheetPart(
      XSSFObjectData objectData) {
    return ExcelDrawingBinarySupport.previewSheetImagePart(objectData).orElse(null);
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow, 0, 0),
        new ExcelDrawingMarker(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }
}
