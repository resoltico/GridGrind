package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import org.junit.jupiter.api.Test;

/** Drawing value-object and authored command coverage. */
class ExcelDrawingValueCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void drawingValueObjectsAndCommandsValidateEveryBranch() {
    ExcelBinaryData data = new ExcelBinaryData(new byte[] {1, 2, 3});
    byte[] extractedBytes = data.bytes();
    extractedBytes[0] = 9;
    assertArrayEquals(new byte[] {1, 2, 3}, data.bytes());
    assertArrayEquals(new byte[] {1, 2, 3}, data.bytes());
    assertEquals(3, data.size());
    assertEquals(new ExcelBinaryData(new byte[] {1, 2, 3}), data);
    assertEquals(new ExcelBinaryData(new byte[] {1, 2, 3}).hashCode(), data.hashCode());
    assertNotEquals(data, new ExcelBinaryData(new byte[] {3, 2, 1}));
    assertNotEquals(data, "payload");
    assertTrue(data.toString().contains("size=3"));
    assertEquals(0, ExcelBinaryData.readback(new byte[0]).size());
    assertThrows(NullPointerException.class, () -> new ExcelBinaryData(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelBinaryData(new byte[0]));

    ExcelDrawingMarker from = new ExcelDrawingMarker(1, 2);
    ExcelDrawingMarker to = new ExcelDrawingMarker(3, 4, 5, 6);
    assertEquals(0, from.dx());
    assertEquals(0, from.dy());
    assertEquals(5, to.dx());
    assertEquals(6, to.dy());
    assertThrows(IllegalArgumentException.class, () -> new ExcelDrawingMarker(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelDrawingMarker(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelDrawingMarker(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelDrawingMarker(0, 0, 0, -1));

    ExcelDrawingAnchor.TwoCell twoCell = new ExcelDrawingAnchor.TwoCell(from, to, null);
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, twoCell.behavior());
    assertThrows(NullPointerException.class, () -> new ExcelDrawingAnchor.TwoCell(null, to, null));
    assertThrows(
        NullPointerException.class, () -> new ExcelDrawingAnchor.TwoCell(from, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2),
                new ExcelDrawingMarker(1, 1),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(2, 1),
                new ExcelDrawingMarker(1, 1),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 0, 5),
                new ExcelDrawingMarker(1, 2, 0, 4),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.TwoCell(
                new ExcelDrawingMarker(1, 2, 5, 0),
                new ExcelDrawingMarker(1, 2, 4, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));

    ExcelDrawingAnchor.OneCell oneCell = new ExcelDrawingAnchor.OneCell(from, 10L, 12L, null);
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, oneCell.behavior());
    assertThrows(
        NullPointerException.class, () -> new ExcelDrawingAnchor.OneCell(null, 10L, 12L, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.OneCell(
                from, 0L, 12L, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingAnchor.OneCell(
                from, 10L, 0L, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));

    ExcelDrawingAnchor.Absolute absolute = new ExcelDrawingAnchor.Absolute(1L, 2L, 3L, 4L, null);
    assertEquals(ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE, absolute.behavior());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDrawingAnchor.Absolute(-1L, 0L, 1L, 1L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDrawingAnchor.Absolute(0L, -1L, 1L, 1L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDrawingAnchor.Absolute(0L, 0L, 0L, 1L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDrawingAnchor.Absolute(0L, 0L, 1L, 0L, null));

    for (ExcelPictureFormat format : ExcelPictureFormat.values()) {
      assertSame(
          format,
          ExcelPicturePoiBridge.fromPoiPictureType(ExcelPicturePoiBridge.toPoiPictureType(format)));
      assertSame(format, ExcelPictureFormat.fromContentType(format.defaultContentType()));
    }
    assertThrows(
        IllegalArgumentException.class, () -> ExcelPicturePoiBridge.fromPoiPictureType(-1));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelPictureFormat.fromContentType("application/x-gridgrind"));

    ExcelPictureDefinition pictureDefinition =
        new ExcelPictureDefinition("Picture", data, ExcelPictureFormat.PNG, twoCell, "preview");
    assertEquals("Picture", pictureDefinition.name());
    assertThrows(
        NullPointerException.class,
        () -> new ExcelPictureDefinition(null, data, ExcelPictureFormat.PNG, twoCell, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPictureDefinition(" ", data, ExcelPictureFormat.PNG, twoCell, null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelPictureDefinition("Picture", null, ExcelPictureFormat.PNG, twoCell, null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelPictureDefinition("Picture", data, null, twoCell, null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelPictureDefinition("Picture", data, ExcelPictureFormat.PNG, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPictureDefinition("Picture", data, ExcelPictureFormat.PNG, twoCell, " "));

    ExcelShapeDefinition defaultShape =
        new ExcelShapeDefinition(
            "Shape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, twoCell, " ", "Queue");
    assertEquals("rect", defaultShape.presetGeometryToken());
    ExcelShapeDefinition connectorShape =
        new ExcelShapeDefinition(
            "Connector", ExcelAuthoredDrawingShapeKind.CONNECTOR, twoCell, null, null);
    assertEquals(ExcelAuthoredDrawingShapeKind.CONNECTOR, connectorShape.kind());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelShapeDefinition(
                null, ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, twoCell, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelShapeDefinition(
                " ", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, twoCell, null, null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelShapeDefinition("Shape", null, twoCell, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelShapeDefinition(
                "Shape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelShapeDefinition(
                "Connector", ExcelAuthoredDrawingShapeKind.CONNECTOR, twoCell, "rect", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelShapeDefinition(
                "Connector", ExcelAuthoredDrawingShapeKind.CONNECTOR, twoCell, null, "text"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelShapeDefinition(
                "Shape", ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, twoCell, "rect", " "));

    ExcelEmbeddedObjectDefinition embeddedDefinition =
        new ExcelEmbeddedObjectDefinition(
            "Embed",
            "Payload",
            "payload.txt",
            "payload.txt",
            data,
            ExcelPictureFormat.PNG,
            new ExcelBinaryData(PNG_PIXEL_BYTES),
            twoCell);
    assertEquals("payload.txt", embeddedDefinition.fileName());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelEmbeddedObjectDefinition(
                " ",
                "Payload",
                "payload.txt",
                "payload.txt",
                data,
                ExcelPictureFormat.PNG,
                data,
                twoCell));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelEmbeddedObjectDefinition(
                null,
                "Payload",
                "payload.txt",
                "payload.txt",
                data,
                ExcelPictureFormat.PNG,
                data,
                twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelEmbeddedObjectDefinition(
                "Embed",
                " ",
                "payload.txt",
                "payload.txt",
                data,
                ExcelPictureFormat.PNG,
                data,
                twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelEmbeddedObjectDefinition(
                "Embed",
                "Payload",
                " ",
                "payload.txt",
                data,
                ExcelPictureFormat.PNG,
                data,
                twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelEmbeddedObjectDefinition(
                "Embed",
                "Payload",
                "payload.txt",
                " ",
                data,
                ExcelPictureFormat.PNG,
                data,
                twoCell));

    assertEquals(
        "Picture",
        new WorkbookDrawingCommand.SetPicture("Ops", pictureDefinition).picture().name());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.SetPicture(null, pictureDefinition));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.SetPicture(" ", pictureDefinition));
    assertThrows(
        NullPointerException.class, () -> new WorkbookDrawingCommand.SetPicture("Ops", null));

    assertEquals("Shape", new WorkbookDrawingCommand.SetShape("Ops", defaultShape).shape().name());
    assertThrows(
        NullPointerException.class, () -> new WorkbookDrawingCommand.SetShape(null, defaultShape));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.SetShape(" ", defaultShape));
    assertThrows(
        NullPointerException.class, () -> new WorkbookDrawingCommand.SetShape("Ops", null));

    assertEquals(
        "Embed",
        new WorkbookDrawingCommand.SetEmbeddedObject("Ops", embeddedDefinition)
            .embeddedObject()
            .name());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.SetEmbeddedObject(null, embeddedDefinition));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.SetEmbeddedObject(" ", embeddedDefinition));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.SetEmbeddedObject("Ops", null));

    assertEquals(
        "Object",
        new WorkbookDrawingCommand.SetDrawingObjectAnchor("Ops", "Object", twoCell).objectName());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.SetDrawingObjectAnchor(null, "Object", twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.SetDrawingObjectAnchor(" ", "Object", twoCell));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.SetDrawingObjectAnchor("Ops", null, twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.SetDrawingObjectAnchor("Ops", " ", twoCell));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.SetDrawingObjectAnchor("Ops", "Object", null));

    assertEquals(
        "Object", new WorkbookDrawingCommand.DeleteDrawingObject("Ops", "Object").objectName());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.DeleteDrawingObject(null, "Object"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.DeleteDrawingObject(" ", "Object"));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingCommand.DeleteDrawingObject("Ops", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.DeleteDrawingObject("Ops", " "));

    ExcelDrawingObjectPayload.Picture picturePayload =
        new ExcelDrawingObjectPayload.Picture(
            "Picture",
            ExcelPictureFormat.PNG,
            "image/png",
            "picture.png",
            "abc123",
            data,
            "preview");
    assertEquals("picture.png", picturePayload.fileName());
    assertNull(
        new ExcelDrawingObjectPayload.Picture(
                "NullDescPicture",
                ExcelPictureFormat.PNG,
                "image/png",
                "picture.png",
                "abc123",
                data,
                null)
            .description());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectPayload.Picture(
                null, ExcelPictureFormat.PNG, "image/png", "picture.png", "abc123", data, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.Picture(
                "Picture", ExcelPictureFormat.PNG, " ", "picture.png", "abc123", data, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectPayload.Picture(
                "Picture", null, "image/png", "picture.png", "abc123", data, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.Picture(
                "Picture", ExcelPictureFormat.PNG, "image/png", " ", "abc123", data, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.Picture(
                "Picture",
                ExcelPictureFormat.PNG,
                "image/png",
                "picture.png",
                "abc123",
                data,
                " "));

    ExcelDrawingObjectPayload.EmbeddedObject embeddedPayload =
        new ExcelDrawingObjectPayload.EmbeddedObject(
            "Embed",
            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
            "application/octet-stream",
            "payload.txt",
            "abc123",
            data,
            "Payload",
            "payload.txt");
    assertEquals("Payload", embeddedPayload.label());
    assertNull(
        new ExcelDrawingObjectPayload.EmbeddedObject(
                "NullOptEmbed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                null,
                "abc123",
                data,
                null,
                null)
            .label());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectPayload.EmbeddedObject(
                "Embed",
                null,
                "application/octet-stream",
                "payload.txt",
                "abc123",
                data,
                "Payload",
                "payload.txt"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.EmbeddedObject(
                "Embed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                " ",
                "abc123",
                data,
                "Payload",
                "payload.txt"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.EmbeddedObject(
                "Embed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                "payload.txt",
                "abc123",
                data,
                " ",
                "payload.txt"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectPayload.EmbeddedObject(
                "Embed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                "payload.txt",
                "abc123",
                data,
                "Payload",
                " "));

    ExcelDrawingObjectSnapshot.Picture pictureSnapshot =
        new ExcelDrawingObjectSnapshot.Picture(
            "Picture",
            twoCell,
            ExcelPictureFormat.PNG,
            "image/png",
            10L,
            "abc123",
            1,
            1,
            "preview");
    assertEquals(10L, pictureSnapshot.byteSize());
    assertNull(
        new ExcelDrawingObjectSnapshot.Picture(
                "NullDescPicSnap",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                null,
                null,
                null)
            .description());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                null,
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture", twoCell, ExcelPictureFormat.PNG, " ", 10L, "abc123", null, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                null,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture", twoCell, null, "image/png", 10L, "abc123", null, null, null));
    assertEquals(
        0L,
        new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                0L,
                "abc123",
                null,
                null,
                null)
            .byteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                -1L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                " ",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                -1,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                null,
                -1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                null,
                null,
                " "));

    ExcelDrawingObjectSnapshot.Shape shapeSnapshot =
        new ExcelDrawingObjectSnapshot.Shape(
            "Shape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, "rect", "Queue", 0);
    assertEquals("rect", shapeSnapshot.presetGeometryToken());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Shape(
                null, twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, null, null, 0));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelDrawingObjectSnapshot.Shape("Shape", twoCell, null, null, null, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Shape(
                "Shape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, " ", null, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Shape(
                "Shape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, null, " ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Shape(
                "Shape", twoCell, ExcelDrawingShapeKind.SIMPLE_SHAPE, null, null, -1));

    ExcelDrawingObjectSnapshot.EmbeddedObject embeddedSnapshot =
        new ExcelDrawingObjectSnapshot.EmbeddedObject(
            "Embed",
            twoCell,
            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
            "Payload",
            "payload.txt",
            "payload.txt",
            "application/octet-stream",
            5L,
            "abc123",
            ExcelPictureFormat.PNG,
            1L,
            "def456");
    assertEquals(1L, embeddedSnapshot.previewByteSize());
    assertNull(
        new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "NullOptEmbedSnap",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                null,
                null,
                null,
                "application/octet-stream",
                5L,
                "abc123",
                null,
                null,
                null)
            .label());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                null,
                null,
                null,
                null,
                "application/octet-stream",
                5L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                " ",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                " ",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                " ",
                "application/octet-stream",
                5L,
                "abc123",
                null,
                null,
                null));
    assertEquals(
        0L,
        new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                0L,
                "abc123",
                null,
                null,
                null)
            .byteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                -1L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                null,
                1L,
                null));
    assertEquals(
        0L,
        new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                ExcelPictureFormat.PNG,
                0L,
                null)
            .previewByteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                ExcelPictureFormat.PNG,
                -1L,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                ExcelPictureFormat.PNG,
                1L,
                " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.EmbeddedObject(
                "Embed",
                twoCell,
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "Payload",
                "payload.txt",
                "payload.txt",
                "application/octet-stream",
                5L,
                "abc123",
                null,
                null,
                "def456"));

    DrawingObjectNotFoundException missing = new DrawingObjectNotFoundException("Ops", "Ghost");
    assertTrue(missing.getMessage().contains("Ghost"));
  }
}
