package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.junit.jupiter.api.Test;

/** Integration tests for drawing, picture, and embedded-object sheet workflows. */
class ExcelDrawingControllerTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void drawingObjectsSupportReadMoveDeleteAndRoundTrip() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-drawing-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  anchor(1, 1, 4, 6),
                  "Queue preview"))
          .setShape(
              new ExcelShapeDefinition(
                  "OpsShape",
                  ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                  anchor(5, 1, 8, 5),
                  "rect",
                  "Queue"))
          .setShape(
              new ExcelShapeDefinition(
                  "OpsConnector",
                  ExcelAuthoredDrawingShapeKind.CONNECTOR,
                  anchor(9, 1, 11, 4),
                  null,
                  null))
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(12, 1, 15, 6)));

      List<ExcelDrawingObjectSnapshot> snapshots = sheet.drawingObjects();
      assertEquals(
          List.of("OpsPicture", "OpsShape", "OpsConnector", "OpsEmbed"),
          snapshots.stream().map(ExcelDrawingObjectSnapshot::name).toList());
      assertEquals(
          ExcelDrawingShapeKind.SIMPLE_SHAPE,
          assertInstanceOf(ExcelDrawingObjectSnapshot.Shape.class, snapshots.get(1)).kind());
      assertEquals(
          ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
          assertInstanceOf(ExcelDrawingObjectSnapshot.EmbeddedObject.class, snapshots.get(3))
              .packagingKind());

      ExcelDrawingObjectPayload.Picture picturePayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.Picture.class, sheet.drawingObjectPayload("OpsPicture"));
      ExcelDrawingObjectPayload.EmbeddedObject embeddedPayload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              sheet.drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(PNG_PIXEL_BYTES, picturePayload.data().bytes());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), embeddedPayload.data().bytes());

      ExcelDrawingAnchor.TwoCell movedAnchor = anchor(6, 2, 10, 7);
      sheet.setDrawingObjectAnchor("OpsShape", movedAnchor);
      ExcelDrawingObjectSnapshot.Shape movedShape =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.Shape.class,
              sheet.drawingObjects().stream()
                  .filter(snapshot -> "OpsShape".equals(snapshot.name()))
                  .findFirst()
                  .orElseThrow());
      assertEquals(movedAnchor, movedShape.anchor());

      sheet.deleteDrawingObject("OpsConnector");
      assertEquals(3, sheet.drawingObjects().size());

      workbook.save(workbookPath);
    }

    try (XSSFWorkbook reopened = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFDrawing drawing = reopened.getSheet("Ops").getDrawingPatriarch();
      assertNotNull(drawing);
      assertEquals(
          List.of("OpsPicture", "OpsShape", "OpsEmbed"),
          drawing.getShapes().stream().map(XSSFShape::getShapeName).toList());
    }
  }

  @Test
  void commentOperationsPreserveRealDrawingObjects() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet commentOnly = workbook.getOrCreateSheet("Comments");
      commentOnly.setComment("A1", new ExcelComment("Review", "GridGrind", false));

      assertTrue(commentOnly.snapshotCell("A1").metadata().comment().isPresent());

      commentOnly.clearComment("A1");
      assertFalse(commentOnly.snapshotCell("A1").metadata().comment().isPresent());

      ExcelSheet withDrawing = workbook.getOrCreateSheet("Ops");
      withDrawing.setPicture(
          new ExcelPictureDefinition(
              "OpsPicture",
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              ExcelPictureFormat.PNG,
              anchor(1, 1, 4, 6),
              "Queue preview"));
      List<ExcelDrawingObjectSnapshot> before = withDrawing.drawingObjects();

      withDrawing.setComment("A1", new ExcelComment("Review", "GridGrind", false));
      assertEquals(before, withDrawing.drawingObjects());

      withDrawing.clearComment("A1");
      assertEquals(before, withDrawing.drawingObjects());
    }
  }

  @Test
  void clearCommentSupportsReopenedPoiCommentWorkbooks() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-clear-comment-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Lead");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      var anchor = drawing.createAnchor(64, 24, 448, 96, 0, 0, 3, 3);
      XSSFComment comment = drawing.createCellComment(anchor);
      comment.setString(new XSSFRichTextString("Review"));
      comment.setAuthor("GridGrind");
      sheet.getRow(0).getCell(0).setCellComment(comment);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      workbook.sheet("Ops").clearComment("A1");
      workbook.save(workbookPath);
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      assertNull(workbook.getSheet("Ops").getRow(0).getCell(0).getCellComment());
    }
  }

  @Test
  void embeddedObjectReadbackFallsBackToDrawingPreviewWhenSheetPreviewMetadataIsMissing()
      throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-gap-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook
          .getOrCreateSheet("Ops")
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(1, 1, 4, 6)));
      workbook.save(workbookPath);
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFObjectData objectData =
          workbook.getSheet("Ops").createDrawingPatriarch().getShapes().stream()
              .filter(XSSFObjectData.class::isInstance)
              .map(XSSFObjectData.class::cast)
              .findFirst()
              .orElseThrow();
      try (XmlCursor cursor = objectData.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjectPayload("OpsEmbed"));

      assertEquals(ExcelPictureFormat.PNG, snapshot.previewFormat());
      assertNotNull(snapshot.previewByteSize());
      assertNotNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  @Test
  void embeddedObjectReadbackSurvivesMissingPreviewImageMetadata() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-embedded-preview-missing-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook
          .getOrCreateSheet("Ops")
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  anchor(1, 1, 4, 6)));
      workbook.save(workbookPath);
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFObjectData objectData =
          workbook.getSheet("Ops").createDrawingPatriarch().getShapes().stream()
              .filter(XSSFObjectData.class::isInstance)
              .map(XSSFObjectData.class::cast)
              .findFirst()
              .orElseThrow();
      try (XmlCursor cursor = objectData.getOleObject().newCursor()) {
        assertTrue(cursor.toChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      objectData.getCTShape().getSpPr().unsetBlipFill();
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelDrawingObjectSnapshot.EmbeddedObject snapshot =
          assertInstanceOf(
              ExcelDrawingObjectSnapshot.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjects().getFirst());
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjectPayload("OpsEmbed"));

      assertNull(snapshot.previewFormat());
      assertNull(snapshot.previewByteSize());
      assertNull(snapshot.previewSha256());
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
    }
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }
}
