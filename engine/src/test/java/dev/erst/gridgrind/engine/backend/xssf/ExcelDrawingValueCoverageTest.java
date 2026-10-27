package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
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
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
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
          format, DrawingAdapter.fromPoiPictureType(DrawingAdapter.toPoiPictureType(format)));
      assertSame(format, ExcelPictureFormat.fromContentType(format.defaultContentType()));
    }
    assertThrows(IllegalArgumentException.class, () -> DrawingAdapter.fromPoiPictureType(-1));
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

    DrawingObjectPayloadReport.Picture picturePayload =
        new DrawingObjectPayloadReport.Picture(
            "Picture",
            ExcelPictureFormat.PNG,
            "image/png",
            "picture.png",
            "abc123",
            java.util.Base64.getEncoder().encodeToString(data.bytes()),
            Optional.of("preview"));
    assertEquals("picture.png", picturePayload.fileName());
    org.junit.jupiter.api.Assertions.assertEquals(
        java.util.Optional.empty(),
        new DrawingObjectPayloadReport.Picture(
                "NullDescPicture",
                ExcelPictureFormat.PNG,
                "image/png",
                "picture.png",
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.empty())
            .description());
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                null,
                ExcelPictureFormat.PNG,
                "image/png",
                "picture.png",
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                "Picture",
                ExcelPictureFormat.PNG,
                " ",
                "picture.png",
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                "Picture",
                null,
                "image/png",
                "picture.png",
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                "Picture",
                ExcelPictureFormat.PNG,
                "image/png",
                " ",
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.Picture(
                "Picture",
                ExcelPictureFormat.PNG,
                "image/png",
                "picture.png",
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.of(" ")));

    DrawingObjectPayloadReport.EmbeddedObject embeddedPayload =
        new DrawingObjectPayloadReport.EmbeddedObject(
            "Embed",
            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
            "application/octet-stream",
            Optional.of("payload.txt"),
            "abc123",
            java.util.Base64.getEncoder().encodeToString(data.bytes()),
            Optional.of("Payload"),
            Optional.of("payload.txt"));
    assertEquals("Payload", embeddedPayload.label().orElseThrow());
    org.junit.jupiter.api.Assertions.assertEquals(
        java.util.Optional.empty(),
        new DrawingObjectPayloadReport.EmbeddedObject(
                "NullOptEmbed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                Optional.empty(),
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.empty(),
                Optional.empty())
            .label());
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "Embed",
                null,
                "application/octet-stream",
                Optional.of("payload.txt"),
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.of("Payload"),
                Optional.of("payload.txt")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "Embed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                Optional.of(" "),
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.of("Payload"),
                Optional.of("payload.txt")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "Embed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                Optional.of("payload.txt"),
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.of(" "),
                Optional.of("payload.txt")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectPayloadReport.EmbeddedObject(
                "Embed",
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                "application/octet-stream",
                Optional.of("payload.txt"),
                "abc123",
                java.util.Base64.getEncoder().encodeToString(data.bytes()),
                Optional.of("Payload"),
                Optional.of(" ")));

    DrawingObjectReport.Picture pictureSnapshot =
        new DrawingObjectReport.Picture(
            "Picture",
            DrawingGeometryFixtures.report(twoCell),
            ExcelPictureFormat.PNG,
            "image/png",
            10L,
            "abc123",
            Optional.of(1),
            Optional.of(1),
            Optional.of("preview"));
    assertEquals(10L, pictureSnapshot.byteSize());
    org.junit.jupiter.api.Assertions.assertEquals(
        java.util.Optional.empty(),
        new DrawingObjectReport.Picture(
                "NullDescPicSnap",
                DrawingGeometryFixtures.report(twoCell),
                ExcelPictureFormat.PNG,
                "image/png",
                10L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
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
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
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
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
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
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
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
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertEquals(
        0L,
        new DrawingObjectReport.Picture(
                "Picture",
                DrawingGeometryFixtures.report(twoCell),
                ExcelPictureFormat.PNG,
                "image/png",
                0L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
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
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
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
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
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
                Optional.of(-1),
                Optional.empty(),
                Optional.empty()));
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
                Optional.empty(),
                Optional.of(-1),
                Optional.empty()));
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
                Optional.empty(),
                Optional.empty(),
                Optional.of(" ")));

    DrawingObjectReport.Shape shapeSnapshot =
        new DrawingObjectReport.Shape(
            "Shape",
            DrawingGeometryFixtures.report(twoCell),
            ExcelDrawingShapeKind.SIMPLE_SHAPE,
            Optional.of("rect"),
            Optional.of("Queue"),
            0);
    assertEquals("rect", shapeSnapshot.presetGeometryToken().orElseThrow());
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectReport.Shape(
                null,
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                Optional.empty(),
                Optional.empty(),
                0));
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                null,
                Optional.empty(),
                Optional.empty(),
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                Optional.of(" "),
                Optional.empty(),
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                Optional.empty(),
                Optional.of(" "),
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.Shape(
                "Shape",
                DrawingGeometryFixtures.report(twoCell),
                ExcelDrawingShapeKind.SIMPLE_SHAPE,
                Optional.empty(),
                Optional.empty(),
                -1));

    DrawingObjectReport.EmbeddedObject embeddedSnapshot =
        new DrawingObjectReport.EmbeddedObject(
            "Embed",
            DrawingGeometryFixtures.report(twoCell),
            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
            Optional.of("Payload"),
            Optional.of("payload.txt"),
            Optional.of("payload.txt"),
            "application/octet-stream",
            5L,
            "abc123",
            Optional.ofNullable(ExcelPictureFormat.PNG),
            Optional.of(1L),
            Optional.of("def456"));
    assertEquals(1L, embeddedSnapshot.previewByteSize().orElseThrow());
    org.junit.jupiter.api.Assertions.assertEquals(
        java.util.Optional.empty(),
        new DrawingObjectReport.EmbeddedObject(
                "NullOptEmbedSnap",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
            .label());
    assertThrows(
        NullPointerException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                null,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of(" "),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of(" "),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of(" "),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertEquals(
        0L,
        new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                0L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
            .byteSize());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                -1L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.of(1L),
                Optional.empty()));
    assertEquals(
        0L,
        new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.ofNullable(ExcelPictureFormat.PNG),
                Optional.of(0L),
                Optional.empty())
            .previewByteSize()
            .orElseThrow());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.ofNullable(ExcelPictureFormat.PNG),
                Optional.of(-1L),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.ofNullable(ExcelPictureFormat.PNG),
                Optional.of(1L),
                Optional.of(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingObjectReport.EmbeddedObject(
                "Embed",
                DrawingGeometryFixtures.report(twoCell),
                ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                Optional.of("Payload"),
                Optional.of("payload.txt"),
                Optional.of("payload.txt"),
                "application/octet-stream",
                5L,
                "abc123",
                Optional.empty(),
                Optional.empty(),
                Optional.of("def456")));

    DrawingObjectNotFoundException missing = new DrawingObjectNotFoundException("Ops", "Ghost");
    assertTrue(missing.getMessage().contains("Ghost"));
  }
}
