package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import java.util.List;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs;

/** Drawing controller public and private branch coverage. */
class ExcelDrawingControllerCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  @SuppressWarnings("PMD.NcssCount")
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
              .filter(ExcelDrawingObjectSnapshot.Chart.class::isInstance)
              .map(ExcelDrawingObjectSnapshot.Chart.class::cast)
              .map(ExcelDrawingObjectSnapshot.Chart::name)
              .findFirst()
              .orElseThrow();
      assertEquals(6, snapshots.size());
      assertEquals(
          1L,
          snapshots.stream().filter(ExcelDrawingObjectSnapshot.Chart.class::isInstance).count());
      assertTrue(
          snapshots.stream()
              .filter(ExcelDrawingObjectSnapshot.Shape.class::isInstance)
              .map(ExcelDrawingObjectSnapshot.Shape.class::cast)
              .anyMatch(snapshot -> snapshot.kind() == ExcelDrawingShapeKind.GROUP));

      assertThrows(
          DrawingObjectNotFoundException.class,
          () -> controller.drawingObjectPayload(sheet, "Missing"));
      assertThrows(
          IllegalArgumentException.class,
          () -> controller.drawingObjectPayload(sheet, simpleShape.getShapeName()));
      controller.setDrawingObjectAnchor(sheet, chartName, anchor(1, 1, 2, 2));
      assertEquals(
          anchor(1, 1, 2, 2),
          controller.drawingObjects(sheet).stream()
              .filter(ExcelDrawingObjectSnapshot.Chart.class::isInstance)
              .map(ExcelDrawingObjectSnapshot.Chart.class::cast)
              .filter(snapshot -> snapshot.name().equals(chartName))
              .findFirst()
              .orElseThrow()
              .anchor());
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
      controller.deleteDrawingObject(sheet, chartName);
      assertTrue(
          controller.drawingObjects(sheet).stream()
              .noneMatch(
                  snapshot ->
                      snapshot instanceof ExcelDrawingObjectSnapshot.Chart chart
                          && chart.name().equals(chartName)));

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
      List<String> remainingNames =
          controller.drawingObjects(sheet).stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .filter(
                  name ->
                      !List.of("OpsEmbed", "OpsPicture", "OpsShape", "OpsConnector").contains(name))
              .toList();
      assertEquals(List.of(), remainingNames);

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
    ExcelBinaryData emptyBinary =
        invoke(controller, "binary", ExcelBinaryData.class, new byte[0], "payload");
    assertEquals(0, emptyBinary.size());

    assertTrue(invoke(controller, "looksLikeOle2Storage", Boolean.class, ole2StorageBytes()));
    assertFalse(invoke(controller, "looksLikeOle2Storage", Boolean.class, PNG_PIXEL_BYTES));
    assertFalse(invoke(controller, "looksLikeOle2Storage", Boolean.class, new byte[0]));
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
}
