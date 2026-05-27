package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingController;
import java.util.Optional;
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
          invoke(
              controller, "previewSheetRelationId", Optional.class, secondObject.getOleObject()));
      assertEquals(
          Optional.empty(),
          invoke(
              controller,
              "previewSheetRelationId",
              Optional.class,
              org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject.Factory
                  .newInstance()));
      assertNotNull(invoke(controller, "previewDrawingRelationId", Optional.class, secondObject));
      String previewSheetRelationId =
          invokeOptional(
                  controller, "previewSheetRelationId", String.class, secondObject.getOleObject())
              .orElseThrow();
      sheet.getPackagePart().removeRelationship(previewSheetRelationId);
      assertNotNull(
          invokeOptional(controller, "previewImagePart", PackagePart.class, secondObject)
              .orElseThrow());

      XSSFObjectData noObjectPr =
          createEmbeddedObject(workbook, drawing, "NoObjectPr", 11, 0, 14, 4);
      try (var cursor = noObjectPr.getOleObject().newCursor()) {
        assertTrue(
            cursor.toChild(
                org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML, "objectPr"));
        cursor.removeXml();
      }
      assertEquals(
          Optional.empty(),
          invoke(controller, "previewSheetRelationId", Optional.class, noObjectPr.getOleObject()));

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
      assertEquals(
          Optional.empty(),
          invoke(
              controller,
              "previewSheetRelationId",
              Optional.class,
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
      assertEquals(
          Optional.empty(),
          invoke(controller, "previewDrawingRelationId", Optional.class, noBlipFill));

      XSSFObjectData noBlip = createEmbeddedObject(workbook, drawing, "NoBlip", 19, 0, 22, 4);
      noBlip.getCTShape().getSpPr().unsetBlipFill();
      noBlip.getCTShape().getSpPr().addNewBlipFill();
      assertEquals(
          Optional.empty(), invoke(controller, "previewDrawingRelationId", Optional.class, noBlip));

      assertEquals(
          Optional.empty(),
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
