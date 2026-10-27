package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
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
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

  @Test
  void repairCopiedEmbeddedObjectsRestoresPoiCloneSheetPackageRelationships() throws IOException {
    SheetCopyAdapter support = new SheetCopyAdapter();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      SheetCopyAdapter.EmbeddedObjectCopySnapshot snapshot = support.snapshot(sourceSheet);
      workbook
          .xssfWorkbook()
          .cloneSheet(workbook.xssfWorkbook().getSheetIndex("Source"), "Replica");

      XSSFSheet replicaPoiSheet = workbook.xssfWorkbook().getSheet("Replica");
      XSSFObjectData copiedObject = requiredEmbeddedObject(replicaPoiSheet, "OpsEmbed");
      assertEquals(Optional.empty(), DrawingAdapter.oleObjectPart(copiedObject));
      assertNull(previewSheetPart(copiedObject));
      assertNull(previewDrawingPart(copiedObject));

      support.repairCopiedEmbeddedObjects(workbook.sheet("Replica"), snapshot);

      XSSFObjectData repairedObject = requiredEmbeddedObject(replicaPoiSheet, "OpsEmbed");
      assertTrue(DrawingAdapter.oleObjectPart(repairedObject).isPresent());
      assertNotNull(previewSheetPart(repairedObject));
      assertNotNull(previewDrawingPart(repairedObject));
      DrawingObjectPayloadReport.EmbeddedObject payload =
          assertInstanceOf(
              DrawingObjectPayloadReport.EmbeddedObject.class,
              workbook.sheet("Replica").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          java.util.Base64.getDecoder().decode(payload.base64Data()));
    }
  }

  @Test
  void copySheetPreservesEmbeddedObjectPayloadsBeforeAndAfterRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-embedded-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("Source");
      sourceSheet.drawings().setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());

      DrawingObjectPayloadReport.EmbeddedObject copiedPayload =
          assertInstanceOf(
              DrawingObjectPayloadReport.EmbeddedObject.class,
              workbook.sheet("Replica").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          java.util.Base64.getDecoder().decode(copiedPayload.base64Data()));
      assertEquals(
          1L,
          workbook.sheet("Replica").drawings().drawingObjects().stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());

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
      DrawingObjectPayloadReport.EmbeddedObject copiedPayload =
          assertInstanceOf(
              DrawingObjectPayloadReport.EmbeddedObject.class,
              reopened.sheet("Replica").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          java.util.Base64.getDecoder().decode(copiedPayload.base64Data()));
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

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sourceSheet = workbook.getOrCreateSheet("LL");
      sourceSheet
          .comments()
          .setComment("C11", new ResolvedComment("Note Name3", "GridGrind", false));
      sourceSheet.drawings().setEmbeddedObject(embeddedObjectDefinition("OpsEmbed", "payload"));

      workbook.sheets().copySheet("LL", "LL_B1", new SheetCopyPosition.AtIndex(0));
      workbook.getOrCreateSheet("LL");
      workbook.getOrCreateSheet("LL");
      workbook.getOrCreateSheet("LL");
      workbook.getOrCreateSheet("LL");

      DrawingObjectPayloadReport.EmbeddedObject copiedPayload =
          assertInstanceOf(
              DrawingObjectPayloadReport.EmbeddedObject.class,
              workbook.sheet("LL_B1").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          java.util.Base64.getDecoder().decode(copiedPayload.base64Data()));
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
      DrawingObjectPayloadReport.EmbeddedObject copiedPayload =
          assertInstanceOf(
              DrawingObjectPayloadReport.EmbeddedObject.class,
              reopened.sheet("LL_B1").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          java.util.Base64.getDecoder().decode(copiedPayload.base64Data()));
      assertEquals(
          1L,
          reopened.sheet("LL_B1").drawings().drawingObjects().stream()
              .filter(snapshot -> "OpsEmbed".equals(snapshot.name()))
              .count());
    }
  }

  private static ResolvedEmbeddedObject embeddedObjectDefinition(
      String objectName, String payloadText) {
    return new ResolvedEmbeddedObject(
        objectName,
        "Payload",
        "payload.txt",
        "payload.txt",
        new BinaryContent(payloadText.getBytes(StandardCharsets.UTF_8)),
        ExcelPictureFormat.PNG,
        new BinaryContent(PNG_PIXEL_BYTES),
        anchor(1, 1, 4, 6));
  }

  private static XSSFObjectData requiredEmbeddedObject(XSSFSheet sheet, String objectName) {
    return sheet.createDrawingPatriarch().getShapes().stream()
        .filter(XSSFObjectData.class::isInstance)
        .map(XSSFObjectData.class::cast)
        .filter(shape -> objectName.equals(DrawingAdapter.resolvedName(shape)))
        .findFirst()
        .orElseThrow();
  }

  private static org.apache.poi.openxml4j.opc.PackagePart previewDrawingPart(
      XSSFObjectData objectData) {
    return DrawingAdapter.previewDrawingRelationId(objectData)
        .flatMap(
            relationId ->
                DrawingAdapter.relatedInternalPart(
                    objectData.getDrawing().getPackagePart(), relationId))
        .orElse(null);
  }

  private static org.apache.poi.openxml4j.opc.PackagePart previewSheetPart(
      XSSFObjectData objectData) {
    return DrawingAdapter.previewSheetImagePart(objectData).orElse(null);
  }

  private static DrawingAnchorSpec.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromColumn, fromRow, 0, 0),
        new DrawingMarkerSpec(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }
}
