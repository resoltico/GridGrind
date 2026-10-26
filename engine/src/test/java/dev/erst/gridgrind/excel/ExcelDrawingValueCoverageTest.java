package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelPicturePoiBridge;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Drawing value-object and authored command coverage. */
class ExcelDrawingValueCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void drawingValueObjectsAndCommandsValidateEveryBranch() {
    BinaryContent data = new BinaryContent(new byte[] {1, 2, 3});

    DrawingMarkerSpec from = new DrawingMarkerSpec(1, 2, 0, 0);
    DrawingMarkerSpec to = new DrawingMarkerSpec(3, 4, 5, 6);
    assertEquals(0, from.dx());
    assertEquals(0, from.dy());
    assertEquals(5, to.dx());
    assertEquals(6, to.dy());
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new DrawingMarkerReport(0, 0, 0, -1));

    DrawingAnchorSpec.TwoCell twoCell =
        new DrawingAnchorSpec.TwoCell(
            from,
            to,
            dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE, twoCell.behavior());
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                null,
                to,
                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                from,
                null,
                dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 0, 0),
                new DrawingMarkerSpec(1, 1, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(2, 1, 0, 0),
                new DrawingMarkerSpec(1, 1, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 0, 5),
                new DrawingMarkerSpec(1, 2, 0, 4),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(1, 2, 5, 0),
                new DrawingMarkerSpec(1, 2, 4, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));

    DrawingAnchorReport.OneCell oneCell =
        new DrawingAnchorReport.OneCell(DrawingGeometryFixtures.marker(from), 10L, 12L, null);
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, oneCell.behavior());
    assertThrows(
        NullPointerException.class, () -> new DrawingAnchorReport.OneCell(null, 10L, 12L, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorReport.OneCell(
                DrawingGeometryFixtures.marker(from),
                0L,
                12L,
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorReport.OneCell(
                DrawingGeometryFixtures.marker(from),
                10L,
                0L,
                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));

    DrawingAnchorReport.Absolute absolute = new DrawingAnchorReport.Absolute(1L, 2L, 3L, 4L, null);
    assertEquals(ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE, absolute.behavior());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(-1L, 0L, 1L, 1L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(0L, -1L, 1L, 1L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(0L, 0L, 0L, 1L, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DrawingAnchorReport.Absolute(0L, 0L, 1L, 0L, null));

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

    ResolvedPicture pictureDefinition =
        new ResolvedPicture(
            "Picture", data, ExcelPictureFormat.PNG, twoCell, Optional.of("preview"));
    assertEquals("Picture", pictureDefinition.name());
    assertThrows(
        NullPointerException.class,
        () -> new ResolvedPicture(null, data, ExcelPictureFormat.PNG, twoCell, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedPicture(" ", data, ExcelPictureFormat.PNG, twoCell, Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ResolvedPicture(
                "Picture", null, ExcelPictureFormat.PNG, twoCell, Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new ResolvedPicture("Picture", data, null, twoCell, Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new ResolvedPicture("Picture", data, ExcelPictureFormat.PNG, null, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedPicture(
                "Picture", data, ExcelPictureFormat.PNG, twoCell, Optional.of(" ")));

    ResolvedShape defaultShape =
        new ResolvedShape.SimpleShape("Shape", twoCell, " rect ", Optional.of("Queue"));
    assertEquals("rect", ((ResolvedShape.SimpleShape) defaultShape).presetGeometryToken());
    ResolvedShape connectorShape = new ShapeSpec.Connector("Connector", twoCell);
    assertEquals(ExcelAuthoredDrawingShapeKind.CONNECTOR, connectorShape.kind());
    assertThrows(
        NullPointerException.class,
        () -> new ResolvedShape.SimpleShape(null, twoCell, "rect", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedShape.SimpleShape(" ", twoCell, "rect", Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new ResolvedShape.SimpleShape("Shape", null, "rect", Optional.empty()));
    assertThrows(NullPointerException.class, () -> new ShapeSpec.Connector("Connector", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedShape.SimpleShape("Shape", twoCell, " ", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedShape.SimpleShape("Shape", twoCell, "rect", Optional.of(" ")));

    ResolvedEmbeddedObject embeddedDefinition =
        new ResolvedEmbeddedObject(
            "Embed",
            "Payload",
            "payload.txt",
            "payload.txt",
            data,
            ExcelPictureFormat.PNG,
            new BinaryContent(PNG_PIXEL_BYTES),
            twoCell);
    assertEquals("payload.txt", embeddedDefinition.fileName());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedEmbeddedObject(
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
            new ResolvedEmbeddedObject(
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
            new ResolvedEmbeddedObject(
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
            new ResolvedEmbeddedObject(
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
            new ResolvedEmbeddedObject(
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

    DrawingObjectReport.Picture pictureSnapshot =
        new DrawingObjectReport.Picture(
            "Picture",
            DrawingGeometryFixtures.report(twoCell),
            ExcelPictureFormat.PNG,
            "image/png",
            10L,
            "abc123",
            1,
            1,
            "preview");
    assertEquals(10L, pictureSnapshot.byteSize());
    assertNull(
        new DrawingObjectReport.Picture(
                "NullDescPicSnap",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                null,
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
                ExcelPictureFormat.PNG,
                " ",
                10L,
                "abc123",
                null,
                null,
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectReport.Picture(
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
                null,
                "image/png",
                10L,
                "abc123",
                null,
                null,
                null));
    assertEquals(
        0L,
        new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                null,
                null,
                " "));

    DrawingObjectReport.Shape shapeSnapshot =
        new DrawingObjectReport.Shape(
            "Shape",
            DrawingGeometryFixtures.report(twoCell),
            ExcelDrawingShapeKind.SIMPLE_SHAPE,
            "rect",
            "Queue",
            0);
    assertEquals("rect", shapeSnapshot.presetGeometryToken());
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectReport.Shape(
                null,
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                null,
                null,
                0));
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape", DrawingGeometryFixtures.report(twoCell), null, null, null, 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                " ",
                null,
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                null,
                " ",
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                null,
                null,
                -1));

    DrawingObjectReport.EmbeddedObject embeddedSnapshot =
        new DrawingObjectReport.EmbeddedObject(
            "Embed",
            DrawingGeometryFixtures.report(twoCell),
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
        new DrawingObjectReport.EmbeddedObject(
                "NullOptEmbedSnap",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
        new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
        new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
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
