package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.security.Provider;
import java.security.Security;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs;

/** Exhaustive coverage tests for the Phase 5 drawing engine surface. */
@SuppressWarnings({
  "PMD.NcssCount",
  "PMD.AvoidAccessibilityAlteration",
  "PMD.SignatureDeclareThrowsException",
  "PMD.CommentRequired"
})
class ExcelDrawingCoverageTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
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
      assertSame(format, ExcelPictureFormat.fromPoiPictureType(format.poiPictureType()));
      assertSame(format, ExcelPictureFormat.fromContentType(format.defaultContentType()));
    }
    assertThrows(IllegalArgumentException.class, () -> ExcelPictureFormat.fromPoiPictureType(-1));
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
        "Picture", new WorkbookCommand.SetPicture("Ops", pictureDefinition).picture().name());
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetPicture(null, pictureDefinition));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetPicture(" ", pictureDefinition));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetPicture("Ops", null));

    assertEquals("Shape", new WorkbookCommand.SetShape("Ops", defaultShape).shape().name());
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetShape(null, defaultShape));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetShape(" ", defaultShape));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.SetShape("Ops", null));

    assertEquals(
        "Embed",
        new WorkbookCommand.SetEmbeddedObject("Ops", embeddedDefinition).embeddedObject().name());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetEmbeddedObject(null, embeddedDefinition));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetEmbeddedObject(" ", embeddedDefinition));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetEmbeddedObject("Ops", null));

    assertEquals(
        "Object",
        new WorkbookCommand.SetDrawingObjectAnchor("Ops", "Object", twoCell).objectName());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetDrawingObjectAnchor(null, "Object", twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetDrawingObjectAnchor(" ", "Object", twoCell));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetDrawingObjectAnchor("Ops", null, twoCell));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetDrawingObjectAnchor("Ops", " ", twoCell));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetDrawingObjectAnchor("Ops", "Object", null));

    assertEquals("Object", new WorkbookCommand.DeleteDrawingObject("Ops", "Object").objectName());
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.DeleteDrawingObject(null, "Object"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.DeleteDrawingObject(" ", "Object"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.DeleteDrawingObject("Ops", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.DeleteDrawingObject("Ops", " "));

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
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.Picture(
                "Picture",
                twoCell,
                ExcelPictureFormat.PNG,
                "image/png",
                0L,
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
                0L,
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
                0L,
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

  @Test
  void drawingControllerCoversPublicAndPrivateBranches() throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertEquals(List.of(), controller.drawingObjects(workbook.createSheet("Blank")));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFPicture picture = createPicture(workbook, drawing, "OpsPicture", 0, 0, 2, 2);
      XSSFSimpleShape simpleShape = createSimpleShape(drawing, "OpsShape", 3, 0, 5, 2);
      XSSFConnector connector = createConnector(drawing, "OpsConnector", 6, 0, 8, 2);
      XSSFShapeGroup group = drawing.createGroup(poiAnchor(drawing, 9, 0, 11, 2));
      drawing.createChart(poiAnchor(drawing, 12, 0, 15, 4));
      XSSFGraphicFrame graphicFrame =
          assertInstanceOf(
              XSSFGraphicFrame.class,
              drawing.getShapes().stream()
                  .filter(XSSFGraphicFrame.class::isInstance)
                  .findFirst()
                  .orElseThrow());
      XSSFObjectData embeddedObject =
          createEmbeddedObject(workbook, drawing, "OpsEmbed", 16, 0, 19, 4);

      List<ExcelDrawingObjectSnapshot> snapshots = controller.drawingObjects(sheet);
      String chartName =
          snapshots.stream()
              .filter(ExcelDrawingObjectSnapshot.Shape.class::isInstance)
              .map(ExcelDrawingObjectSnapshot.Shape.class::cast)
              .filter(snapshot -> snapshot.kind() == ExcelDrawingShapeKind.GRAPHIC_FRAME)
              .map(ExcelDrawingObjectSnapshot.Shape::name)
              .findFirst()
              .orElseThrow();
      assertEquals(6, snapshots.size());
      assertTrue(
          snapshots.stream()
              .filter(ExcelDrawingObjectSnapshot.Shape.class::isInstance)
              .map(ExcelDrawingObjectSnapshot.Shape.class::cast)
              .anyMatch(snapshot -> snapshot.kind() == ExcelDrawingShapeKind.GROUP));
      assertTrue(
          snapshots.stream()
              .filter(ExcelDrawingObjectSnapshot.Shape.class::isInstance)
              .map(ExcelDrawingObjectSnapshot.Shape.class::cast)
              .anyMatch(snapshot -> snapshot.kind() == ExcelDrawingShapeKind.GRAPHIC_FRAME));

      assertThrows(
          DrawingObjectNotFoundException.class,
          () -> controller.drawingObjectPayload(sheet, "Missing"));
      assertThrows(
          IllegalArgumentException.class,
          () -> controller.drawingObjectPayload(sheet, simpleShape.getShapeName()));
      assertThrows(
          IllegalArgumentException.class,
          () -> controller.setDrawingObjectAnchor(sheet, chartName, anchor(1, 1, 2, 2)));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setShape(
                  sheet,
                  new ExcelShapeDefinition(
                      "OpsBrokenShape",
                      ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                      anchor(1, 1, 2, 2),
                      "invalid-shape",
                      null)));

      String pictureDefaultName = invoke(controller, "defaultName", String.class, picture);
      String connectorDefaultName = invoke(controller, "defaultName", String.class, connector);
      String groupDefaultName = invoke(controller, "defaultName", String.class, group);
      String graphicFrameDefaultName =
          invoke(controller, "defaultName", String.class, graphicFrame);
      String simpleDefaultName = invoke(controller, "defaultName", String.class, simpleShape);
      String objectDefaultName = invoke(controller, "defaultName", String.class, embeddedObject);
      assertTrue(pictureDefaultName.startsWith("Picture-"));
      assertTrue(connectorDefaultName.startsWith("Connector-"));
      assertTrue(groupDefaultName.startsWith("Group-"));
      assertTrue(graphicFrameDefaultName.startsWith("GraphicFrame-"));
      assertTrue(simpleDefaultName.startsWith("Shape-"));
      assertTrue(objectDefaultName.startsWith("Object-"));

      assertNotNull(invoke(controller, "shapeXml", Object.class, picture));
      assertNotNull(invoke(controller, "shapeXml", Object.class, connector));
      assertNotNull(invoke(controller, "shapeXml", Object.class, group));
      assertNotNull(invoke(controller, "shapeXml", Object.class, graphicFrame));
      assertNotNull(invoke(controller, "shapeXml", Object.class, simpleShape));
      assertNotNull(invoke(controller, "shapeXml", Object.class, embeddedObject));
      simpleShape.getCTShape().getSpPr().unsetPrstGeom();
      simpleShape.getCTShape().getNvSpPr().getCNvPr().setName("");
      assertNull(
          invoke(controller, "snapshotShape", ExcelDrawingObjectSnapshot.Shape.class, simpleShape)
              .presetGeometryToken());
      assertTrue(
          invoke(controller, "resolvedName", String.class, simpleShape).startsWith("Shape-"));
      assertThrows(
          IllegalArgumentException.class, () -> controller.deleteDrawingObject(sheet, chartName));

      simpleShape.getCTShape().getNvSpPr().getCNvPr().setName("Duplicate");
      connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName("Duplicate");
      assertThrows(
          IllegalArgumentException.class, () -> controller.deleteDrawingObject(sheet, "Duplicate"));
      simpleShape.getCTShape().getNvSpPr().getCNvPr().setName("OpsShape");
      connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName("OpsConnector");
      invokeVoid(
          controller,
          "removeOleObject",
          sheet,
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject.Factory.newInstance());

      controller.deleteDrawingObject(sheet, embeddedObject.getShapeName());
      controller.deleteDrawingObject(sheet, picture.getShapeName());
      controller.deleteDrawingObject(sheet, simpleShape.getShapeName());
      controller.deleteDrawingObject(sheet, connector.getShapeName());
      controller.deleteDrawingObject(sheet, group.getShapeName());
      assertEquals(
          2,
          controller.drawingObjects(sheet).stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .filter(
                  name ->
                      !List.of("OpsEmbed", "OpsPicture", "OpsShape", "OpsConnector").contains(name))
              .count());

      IllegalStateException unsupportedSnapshot =
          assertInvocationFailure(
              IllegalStateException.class,
              () -> invoke(controller, "snapshot", Object.class, drawing, new UnsupportedShape()));
      assertTrue(unsupportedSnapshot.getMessage().contains("Unsupported drawing shape type"));

      IllegalStateException unsupportedShapeXml =
          assertInvocationFailure(
              IllegalStateException.class,
              () -> invoke(controller, "shapeXml", Object.class, new UnsupportedShape()));
      assertTrue(unsupportedShapeXml.getMessage().contains("Unsupported drawing shape type"));

      IllegalStateException unsupportedDefaultName =
          assertInvocationFailure(
              IllegalStateException.class,
              () -> invoke(controller, "defaultName", String.class, new UnsupportedShape()));
      assertTrue(unsupportedDefaultName.getMessage().contains("Unsupported drawing shape type"));
    }

    CTShape oneCellShape = oneCellShape();
    ExcelDrawingAnchor.OneCell reflectedOneCell =
        assertInstanceOf(
            ExcelDrawingAnchor.OneCell.class,
            invoke(controller, "snapshotAnchor", ExcelDrawingAnchor.class, oneCellShape));
    assertEquals(ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE, reflectedOneCell.behavior());

    CTShape absoluteShape = absoluteShape();
    ExcelDrawingAnchor.Absolute reflectedAbsolute =
        assertInstanceOf(
            ExcelDrawingAnchor.Absolute.class,
            invoke(controller, "snapshotAnchor", ExcelDrawingAnchor.class, absoluteShape));
    assertEquals(ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE, reflectedAbsolute.behavior());

    IllegalStateException missingParent =
        assertInvocationFailure(
            IllegalStateException.class,
            () ->
                invoke(
                    controller,
                    "snapshotAnchor",
                    ExcelDrawingAnchor.class,
                    CTShape.Factory.newInstance()));
    assertTrue(missingParent.getMessage().contains("missing its parent anchor"));

    IllegalStateException unsupportedParent =
        assertInvocationFailure(
            IllegalStateException.class,
            () ->
                invoke(
                    controller,
                    "snapshotAnchor",
                    ExcelDrawingAnchor.class,
                    unsupportedParentShape()));
    assertTrue(unsupportedParent.getMessage().contains("Unsupported parent anchor type"));

    CTDrawing ctDrawing = CTDrawing.Factory.newInstance();
    CTTwoCellAnchor twoCellAnchor = ctDrawing.addNewTwoCellAnchor();
    populateTwoCellAnchor(twoCellAnchor);
    CTOneCellAnchor oneCellAnchor = ctDrawing.addNewOneCellAnchor();
    populateOneCellAnchor(oneCellAnchor);
    CTOneCellAnchor secondOneCellAnchor = ctDrawing.addNewOneCellAnchor();
    populateOneCellAnchor(secondOneCellAnchor);
    CTAbsoluteAnchor absoluteAnchor = ctDrawing.addNewAbsoluteAnchor();
    populateAbsoluteAnchor(absoluteAnchor);
    CTAbsoluteAnchor secondAbsoluteAnchor = ctDrawing.addNewAbsoluteAnchor();
    populateAbsoluteAnchor(secondAbsoluteAnchor);
    invokeVoid(controller, "removeTwoCellAnchor", ctDrawing, twoCellAnchor);
    invokeVoid(controller, "removeOneCellAnchor", ctDrawing, secondOneCellAnchor);
    invokeVoid(controller, "removeAbsoluteAnchor", ctDrawing, secondAbsoluteAnchor);
    assertEquals(0, ctDrawing.sizeOfTwoCellAnchorArray());
    assertEquals(1, ctDrawing.sizeOfOneCellAnchorArray());
    assertEquals(1, ctDrawing.sizeOfAbsoluteAnchorArray());
    invokeVoid(controller, "removeOneCellAnchor", ctDrawing, oneCellAnchor);
    invokeVoid(controller, "removeAbsoluteAnchor", ctDrawing, absoluteAnchor);
    assertEquals(0, ctDrawing.sizeOfOneCellAnchorArray());
    assertEquals(0, ctDrawing.sizeOfAbsoluteAnchorArray());
    assertInvocationFailure(
        IllegalStateException.class,
        () -> invokeVoid(controller, "removeTwoCellAnchor", ctDrawing, twoCellAnchor));
    assertInvocationFailure(
        IllegalStateException.class,
        () -> invokeVoid(controller, "removeOneCellAnchor", ctDrawing, oneCellAnchor));
    assertInvocationFailure(
        IllegalStateException.class,
        () -> invokeVoid(controller, "removeAbsoluteAnchor", ctDrawing, absoluteAnchor));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      CTDrawing liveDrawing = drawing.getCTDrawing();
      CTTwoCellAnchor liveTwoCellAnchor = liveDrawing.addNewTwoCellAnchor();
      populateTwoCellAnchor(liveTwoCellAnchor);
      CTOneCellAnchor liveOneCellAnchor = liveDrawing.addNewOneCellAnchor();
      populateOneCellAnchor(liveOneCellAnchor);
      CTAbsoluteAnchor liveAbsoluteAnchor = liveDrawing.addNewAbsoluteAnchor();
      populateAbsoluteAnchor(liveAbsoluteAnchor);
      invokeVoid(controller, "removeParentAnchor", drawing, liveTwoCellAnchor);
      invokeVoid(controller, "removeParentAnchor", drawing, liveOneCellAnchor);
      invokeVoid(controller, "removeParentAnchor", drawing, liveAbsoluteAnchor);
      IllegalStateException nullParent =
          assertInvocationFailure(
              IllegalStateException.class,
              () -> invokeVoid(controller, "removeParentAnchor", drawing, null));
      assertTrue(nullParent.getMessage().contains("missing its parent anchor"));
      IllegalStateException unsupportedAnchorType =
          assertInvocationFailure(
              IllegalStateException.class,
              () ->
                  invokeVoid(controller, "removeParentAnchor", drawing, unsupportedParentShape()));
      assertTrue(unsupportedAnchorType.getMessage().contains("Unsupported parent anchor type"));

      IllegalArgumentException wrongAnchorType =
          assertInvocationFailure(
              IllegalArgumentException.class,
              () ->
                  invokeVoid(
                      controller,
                      "updateAnchorInPlace",
                      sheet,
                      "Shape",
                      oneCellAnchor,
                      anchor(1, 1, 2, 2)));
      assertTrue(wrongAnchorType.getMessage().contains("is not backed by a two-cell anchor"));
    }

    assertEquals(
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE,
        invoke(controller, "behavior", ExcelDrawingAnchorBehavior.class, (Object) null));
    assertEquals(
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE,
        invoke(controller, "behavior", ExcelDrawingAnchorBehavior.class, STEditAs.ONE_CELL));
    assertEquals(
        ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE,
        invoke(controller, "behavior", ExcelDrawingAnchorBehavior.class, STEditAs.ABSOLUTE));
    assertEquals(
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE,
        invoke(controller, "behavior", ExcelDrawingAnchorBehavior.class, STEditAs.TWO_CELL));
    assertEquals(
        "value", invoke(controller, "requireNonBlank", String.class, "value", "fieldName"));
    assertInvocationFailure(
        NullPointerException.class,
        () -> invoke(controller, "requireNonBlank", String.class, null, "fieldName"));
    assertInvocationFailure(
        IllegalArgumentException.class,
        () -> invoke(controller, "requireNonBlank", String.class, " ", "fieldName"));
    assertEquals("first", invoke(controller, "firstNonBlank", String.class, "first", "second"));
    assertEquals("second", invoke(controller, "firstNonBlank", String.class, " ", "second"));
    assertNull(invoke(controller, "firstNonBlank", String.class, " ", " "));
    assertEquals(ShapeTypes.RECT, invoke(controller, "shapeType", Integer.class, "rect"));
    assertInvocationFailure(
        IllegalArgumentException.class,
        () -> invoke(controller, "shapeType", Integer.class, "invalid-shape"));
    assertNotNull(
        invoke(
            controller, "toPoiBehavior", Object.class, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertNotNull(
        invoke(
            controller,
            "toPoiBehavior",
            Object.class,
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));
    assertNotNull(
        invoke(
            controller,
            "toPoiBehavior",
            Object.class,
            ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE));
    assertEquals(
        STEditAs.TWO_CELL,
        invoke(
            controller,
            "toPoiEditAs",
            STEditAs.Enum.class,
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
    assertEquals(
        STEditAs.ONE_CELL,
        invoke(
            controller,
            "toPoiEditAs",
            STEditAs.Enum.class,
            ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE));
    assertEquals(
        STEditAs.ABSOLUTE,
        invoke(
            controller,
            "toPoiEditAs",
            STEditAs.Enum.class,
            ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE));

    ExcelBinaryData binary =
        invoke(controller, "binary", ExcelBinaryData.class, PNG_PIXEL_BYTES, "image");
    assertEquals(PNG_PIXEL_BYTES.length, binary.size());
    IllegalStateException missingBytes =
        assertInvocationFailure(
            IllegalStateException.class,
            () -> invoke(controller, "binary", ExcelBinaryData.class, new byte[0], "payload"));
    assertTrue(missingBytes.getMessage().contains("Missing payload bytes"));

    assertTrue(invoke(controller, "looksLikeOle2Storage", Boolean.class, ole2StorageBytes()));
    assertFalse(invoke(controller, "looksLikeOle2Storage", Boolean.class, PNG_PIXEL_BYTES));
    assertEquals(
        "picture.png",
        invoke(
            controller,
            "partFileName",
            String.class,
            new FixedBytesPackagePart("/xl/media/picture.png", "image/png", PNG_PIXEL_BYTES)));
    assertEquals(
        "5be3713aa69589bb763cc4949206c21415737e47808e8646871b85e671c947d2",
        invoke(controller, "sha256", String.class, PNG_PIXEL_BYTES));

    assertArrayEquals(
        PNG_PIXEL_BYTES,
        invoke(
            controller,
            "partBytes",
            byte[].class,
            new FixedBytesPackagePart("/xl/media/picture.png", "image/png", PNG_PIXEL_BYTES)));
    IllegalStateException readFailure =
        assertInvocationFailure(
            IllegalStateException.class,
            () ->
                invoke(
                    controller,
                    "partBytes",
                    byte[].class,
                    new FailingPackagePart("/xl/media/picture.png", "image/png")));
    assertTrue(readFailure.getMessage().contains("Failed to read package part bytes"));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFObjectData objectData = createEmbeddedObject(workbook, drawing, "OpsEmbed", 1, 1, 3, 3);
      assertNotNull(
          invoke(
              controller,
              "relatedInternalPart",
              Object.class,
              sheet.getPackagePart(),
              objectData.getOleObject().getId()));
      assertNull(
          invoke(controller, "relatedInternalPart", Object.class, sheet.getPackagePart(), null));
      assertNull(
          invoke(controller, "relatedInternalPart", Object.class, sheet.getPackagePart(), " "));
      assertNull(
          invoke(
              controller, "relatedInternalPart", Object.class, sheet.getPackagePart(), "missing"));
      sheet
          .getPackagePart()
          .addExternalRelationship(
              "https://example.com/resource", "urn:gridgrind:test", "rIdExternal");
      assertNull(
          invoke(
              controller,
              "relatedInternalPart",
              Object.class,
              sheet.getPackagePart(),
              "rIdExternal"));
      IllegalStateException invalidRelatedPart =
          assertInvocationFailure(
              IllegalStateException.class,
              () ->
                  invoke(
                      controller,
                      "relatedInternalPart",
                      Object.class,
                      new InvalidTargetPackagePart(
                          "/xl/worksheets/sheet1.xml", "application/vnd.ms-excel.sheet"),
                      "rId1"));
      assertTrue(
          invalidRelatedPart.getMessage().contains("Failed to resolve related package part"));
      invokeVoid(controller, "cleanupWorkbookImagePartIfUnused", workbook, null);
      assertFalse(
          invoke(
              controller,
              "imagePartUsed",
              Boolean.class,
              workbook,
              PackagingURIHelper.createPartName("/xl/media/unused.png")));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Blank");
      assertFalse(
          invoke(
              controller,
              "imagePartUsed",
              Boolean.class,
              workbook,
              PackagingURIHelper.createPartName("/xl/media/unused.png")));
      invokeVoid(
          controller,
          "removeOleObject",
          workbook.getSheet("Blank"),
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject.Factory.newInstance());
    }

    IllegalStateException invalidRelationshipsFailure =
        assertInvocationFailure(
            IllegalStateException.class,
            () ->
                invokeVoid(
                    controller,
                    "removeRelationshipsToPart",
                    new InvalidRelationshipsPackagePart(
                        "/xl/worksheets/sheet1.xml", "application/vnd.ms-excel.sheet"),
                    PackagingURIHelper.createPartName("/xl/media/unused.png")));
    assertTrue(
        invalidRelationshipsFailure
            .getMessage()
            .contains("Failed to inspect package relationships"));

    invokeVoid(
        controller,
        "removeRelationshipsToPart",
        new RelationshipPartPackagePart(
            "/xl/_rels/workbook.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml"),
        PackagingURIHelper.createPartName("/xl/media/unused.png"));
  }

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
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
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

    try (OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream())) {
      java.lang.reflect.Method addPackagePart =
          OPCPackage.class.getDeclaredMethod("addPackagePart", PackagePart.class);
      addPackagePart.setAccessible(true);
      addPackagePart.invoke(
          pkg,
          new InvalidRelationshipsPackagePart(
              "/xl/worksheets/sheet1.xml",
              "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
      IllegalStateException invalidPackageFailure =
          assertInvocationFailure(
              IllegalStateException.class,
              () ->
                  invokeVoid(
                      controller,
                      "cleanupPackagePartIfUnused",
                      pkg,
                      PackagingURIHelper.createPartName("/xl/media/unused.png")));
      assertTrue(
          invalidPackageFailure.getMessage().contains("Failed to inspect package relationships"));
    }

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

  private static XSSFPicture createPicture(
      XSSFWorkbook workbook,
      XSSFDrawing drawing,
      String name,
      int col1,
      int row1,
      int col2,
      int row2) {
    int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
    XSSFPicture picture =
        drawing.createPicture(poiAnchor(drawing, col1, row1, col2, row2), pictureIndex);
    picture.getCTPicture().getNvPicPr().getCNvPr().setName(name);
    return picture;
  }

  private static XSSFSimpleShape createSimpleShape(
      XSSFDrawing drawing, String name, int col1, int row1, int col2, int row2) {
    XSSFSimpleShape shape = drawing.createSimpleShape(poiAnchor(drawing, col1, row1, col2, row2));
    shape.getCTShape().getNvSpPr().getCNvPr().setName(name);
    return shape;
  }

  private static XSSFConnector createConnector(
      XSSFDrawing drawing, String name, int col1, int row1, int col2, int row2) {
    XSSFConnector connector = drawing.createConnector(poiAnchor(drawing, col1, row1, col2, row2));
    connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(name);
    return connector;
  }

  private static XSSFObjectData createEmbeddedObject(
      XSSFWorkbook workbook,
      XSSFDrawing drawing,
      String name,
      int col1,
      int row1,
      int col2,
      int row2)
      throws IOException {
    int storageId =
        workbook.addOlePackage(
            "payload".getBytes(StandardCharsets.UTF_8), "Payload", "payload.txt", "payload.txt");
    int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
    XSSFObjectData objectData =
        drawing.createObjectData(
            poiAnchor(drawing, col1, row1, col2, row2), storageId, pictureIndex);
    objectData.getCTShape().getNvSpPr().getCNvPr().setName(name);
    return objectData;
  }

  private static XSSFClientAnchor poiAnchor(
      XSSFDrawing drawing, int col1, int row1, int col2, int row2) {
    return drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static CTShape oneCellShape() {
    CTOneCellAnchor anchor = CTOneCellAnchor.Factory.newInstance();
    populateOneCellAnchor(anchor);
    return anchor.getSp();
  }

  private static CTShape absoluteShape() {
    CTAbsoluteAnchor anchor = CTAbsoluteAnchor.Factory.newInstance();
    populateAbsoluteAnchor(anchor);
    return anchor.getSp();
  }

  private static CTShape unsupportedParentShape() {
    return org.openxmlformats
        .schemas
        .drawingml
        .x2006
        .spreadsheetDrawing
        .CTGroupShape
        .Factory
        .newInstance()
        .addNewSp();
  }

  private static void populateTwoCellAnchor(CTTwoCellAnchor anchor) {
    CTMarker from = anchor.addNewFrom();
    from.setCol(1);
    from.setRow(2);
    from.setColOff(3L);
    from.setRowOff(4L);
    CTMarker to = anchor.addNewTo();
    to.setCol(5);
    to.setRow(6);
    to.setColOff(7L);
    to.setRowOff(8L);
    anchor.addNewSp();
  }

  private static void populateOneCellAnchor(CTOneCellAnchor anchor) {
    CTMarker from = anchor.addNewFrom();
    from.setCol(1);
    from.setRow(2);
    from.setColOff(3L);
    from.setRowOff(4L);
    anchor.addNewExt().setCx(5L);
    anchor.getExt().setCy(6L);
    anchor.addNewSp();
  }

  private static void populateAbsoluteAnchor(CTAbsoluteAnchor anchor) {
    anchor.addNewPos().setX(7L);
    anchor.getPos().setY(8L);
    anchor.addNewExt().setCx(9L);
    anchor.getExt().setCy(10L);
    anchor.addNewSp();
  }

  private static byte[] ole2StorageBytes() throws IOException {
    try (var filesystem = new org.apache.poi.poifs.filesystem.POIFSFileSystem();
        var output = new ByteArrayOutputStream()) {
      filesystem.createDirectory("Root");
      filesystem.writeFilesystem(output);
      return output.toByteArray();
    }
  }

  private static <T> T invoke(Object target, String name, Class<T> returnType, Object... args)
      throws Exception {
    Class<?>[] parameterTypes = new Class<?>[args.length];
    for (int index = 0; index < args.length; index++) {
      parameterTypes[index] = args[index] == null ? Object.class : args[index].getClass();
    }
    Method method = findMethod(target.getClass(), name, args);
    method.setAccessible(true);
    return returnType.cast(method.invoke(target, args));
  }

  private static void invokeVoid(Object target, String name, Object... args) throws Exception {
    Method method = findMethod(target.getClass(), name, args);
    method.setAccessible(true);
    method.invoke(target, args);
  }

  private static Method findMethod(Class<?> type, String name, Object[] args)
      throws NoSuchMethodException {
    List<Method> matches = new ArrayList<>();
    for (Method method : type.getDeclaredMethods()) {
      if (!method.getName().equals(name) || method.getParameterCount() != args.length) {
        continue;
      }
      boolean compatible = true;
      Class<?>[] parameterTypes = method.getParameterTypes();
      for (int index = 0; index < args.length; index++) {
        if (args[index] == null) {
          continue;
        }
        if (!wrap(parameterTypes[index]).isAssignableFrom(args[index].getClass())) {
          compatible = false;
          break;
        }
      }
      if (compatible) {
        matches.add(method);
      }
    }
    if (matches.size() != 1) {
      throw new NoSuchMethodException(name + " with " + args.length + " parameters");
    }
    return matches.getFirst();
  }

  private static Class<?> wrap(Class<?> type) {
    if (!type.isPrimitive()) {
      return type;
    }
    return switch (type.getName()) {
      case "boolean" -> Boolean.class;
      case "byte" -> Byte.class;
      case "char" -> Character.class;
      case "double" -> Double.class;
      case "float" -> Float.class;
      case "int" -> Integer.class;
      case "long" -> Long.class;
      case "short" -> Short.class;
      default -> type;
    };
  }

  private static <T extends Throwable> T assertInvocationFailure(
      Class<T> type, ThrowingRunnable runnable) {
    InvocationTargetException failure =
        assertThrows(InvocationTargetException.class, runnable::run);
    return assertInstanceOf(type, failure.getCause());
  }

  @FunctionalInterface
  private interface ThrowingRunnable {
    void run() throws Exception;
  }

  private static final class UnsupportedShape extends XSSFShape {
    @Override
    public String getShapeName() {
      return "Unsupported";
    }

    @Override
    protected CTShapeProperties getShapeProperties() {
      return null;
    }
  }

  private abstract static class TestPackagePart extends PackagePart {
    private final OPCPackage container;

    private TestPackagePart(String partName, String contentType) throws InvalidFormatException {
      this(OPCPackage.create(new ByteArrayOutputStream()), partName, contentType);
    }

    private TestPackagePart(OPCPackage container, String partName, String contentType)
        throws InvalidFormatException {
      super(container, PackagingURIHelper.createPartName(partName), contentType);
      this.container = container;
    }

    @Override
    protected OutputStream getOutputStreamImpl() {
      return new ByteArrayOutputStream();
    }

    @Override
    public boolean save(OutputStream outputStream) {
      return true;
    }

    @Override
    public boolean load(InputStream inputStream) {
      return true;
    }

    @Override
    public void close() {
      try {
        container.close();
      } catch (IOException exception) {
        throw new IllegalStateException("Failed to close test package", exception);
      }
    }

    @Override
    public void flush() {}
  }

  private static final class FixedBytesPackagePart extends TestPackagePart {
    private final byte[] bytes;

    private FixedBytesPackagePart(String partName, String contentType, byte[] bytes)
        throws InvalidFormatException {
      super(partName, contentType);
      this.bytes = bytes.clone();
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(bytes);
    }
  }

  private static final class FailingPackagePart extends TestPackagePart {
    private FailingPackagePart(String partName, String contentType) throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    protected InputStream getInputStreamImpl() throws IOException {
      throw new IOException("broken");
    }
  }

  private static final class InvalidTargetPackagePart extends TestPackagePart {
    private InvalidTargetPackagePart(String partName, String contentType)
        throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    public PackageRelationship getRelationship(String relationshipId) {
      return new PackageRelationship(
          getPackage(),
          this,
          java.net.URI.create(".."),
          TargetMode.INTERNAL,
          "urn:test",
          relationshipId);
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(new byte[0]);
    }
  }

  private static final class InvalidRelationshipsPackagePart extends TestPackagePart {
    private InvalidRelationshipsPackagePart(String partName, String contentType)
        throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    public PackageRelationshipCollection getRelationships() throws InvalidFormatException {
      throw new InvalidFormatException("broken relationships");
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(new byte[0]);
    }
  }

  private static final class RelationshipPartPackagePart extends TestPackagePart {
    private RelationshipPartPackagePart(String partName, String contentType)
        throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    public boolean isRelationshipPart() {
      return true;
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(new byte[0]);
    }
  }
}
