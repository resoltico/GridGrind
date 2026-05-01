package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Drawing embedded-preview helper coverage. */
class ExcelDrawingEmbeddedPreviewCoverageTest extends ExcelDrawingCoverageTestSupport {
  @Test
  void drawingControllerReflectiveEmbeddedPreviewHelpersCoverRemainingBranches() throws Exception {
    ExcelDrawingController controller = new ExcelDrawingController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFPicture picture = createPicture(workbook, drawing, "OpsPicture", 0, 0, 2, 2);
      XSSFObjectData firstObject =
          createEmbeddedObject(workbook, drawing, "FirstEmbed", 3, 0, 6, 4);
      XSSFObjectData secondObject =
          createEmbeddedObject(workbook, drawing, "SecondEmbed", 7, 0, 10, 4);

      assertFalse(
          invoke(
              controller,
              "imagePartUsed",
              Boolean.class,
              workbook,
              PackagingURIHelper.createPartName("/xl/media/missing-gridgrind.png")));

      invokeVoid(controller, "cleanupWorkbookImagePartIfUnused", workbook, null);
      invokeVoid(
          controller,
          "cleanupWorkbookImagePartIfUnused",
          workbook,
          picture.getPictureData().getPackagePart().getPartName());
      assertTrue(
          workbook
              .getPackage()
              .containPart(picture.getPictureData().getPackagePart().getPartName()));

      assertNotNull(
          invoke(controller, "previewSheetRelationId", String.class, secondObject.getOleObject()));
      assertNull(
          invoke(
              controller,
              "previewSheetRelationId",
              String.class,
              org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject.Factory
                  .newInstance()));
      assertNotNull(invoke(controller, "previewDrawingRelationId", String.class, secondObject));
      String previewSheetRelationId =
          invoke(controller, "previewSheetRelationId", String.class, secondObject.getOleObject());
      sheet.getPackagePart().removeRelationship(previewSheetRelationId);
      assertNotNull(invoke(controller, "previewImagePart", PackagePart.class, secondObject));

      XSSFObjectData noObjectPr =
          createEmbeddedObject(workbook, drawing, "NoObjectPr", 11, 0, 14, 4);
      try (var cursor = noObjectPr.getOleObject().newCursor()) {
        assertTrue(
            cursor.toChild(
                org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      assertNull(
          invoke(controller, "previewSheetRelationId", String.class, noObjectPr.getOleObject()));

      XSSFObjectData noPreviewAttribute =
          createEmbeddedObject(workbook, drawing, "NoPreviewAttribute", 11, 5, 14, 9);
      try (var cursor = noPreviewAttribute.getOleObject().newCursor()) {
        assertTrue(
            cursor.toChild(
                org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeAttribute(
            new javax.xml.namespace.QName(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id", "r"));
      }
      assertNull(
          invoke(
              controller,
              "previewSheetRelationId",
              String.class,
              noPreviewAttribute.getOleObject()));

      noObjectPr.getCTShape().getSpPr().unsetBlipFill();
      assertFalse(
          invoke(
              controller,
              "imagePartUsed",
              Boolean.class,
              workbook,
              PackagingURIHelper.createPartName("/xl/media/still-missing-gridgrind.png")));

      int oleObjectsBeforeRemoval = sheet.getCTWorksheet().getOleObjects().sizeOfOleObjectArray();
      invokeVoid(controller, "removeOleObject", sheet, firstObject.getOleObject());
      assertTrue(sheet.getCTWorksheet().isSetOleObjects());
      assertEquals(
          oleObjectsBeforeRemoval - 1,
          sheet.getCTWorksheet().getOleObjects().sizeOfOleObjectArray());

      XSSFObjectData noBlipFill =
          createEmbeddedObject(workbook, drawing, "NoBlipFill", 15, 0, 18, 4);
      noBlipFill.getCTShape().getSpPr().unsetBlipFill();
      assertNull(invoke(controller, "previewDrawingRelationId", String.class, noBlipFill));

      XSSFObjectData noBlip = createEmbeddedObject(workbook, drawing, "NoBlip", 19, 0, 22, 4);
      noBlip.getCTShape().getSpPr().unsetBlipFill();
      noBlip.getCTShape().getSpPr().addNewBlipFill();
      assertNull(invoke(controller, "previewDrawingRelationId", String.class, noBlip));

      assertNull(
          invoke(
              controller,
              "parentAnchor",
              Object.class,
              org.apache.xmlbeans.XmlObject.Factory.newInstance()));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData noOleId =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "NoOleId", 27, 0, 30, 4);
      noOleId.getOleObject().unsetId();
      controller.deleteDrawingObject(sheet, noOleId.getShapeName());
      assertTrue(
          controller.drawingObjects(sheet).stream()
              .noneMatch(snapshot -> "NoOleId".equals(snapshot.name())));
    }
  }
}
