package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.microsoft.schemas.office.excel.CTClientData;
import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.junit.jupiter.api.Test;

/** Branch coverage for VML-backed signature-line inspection and mutation. */
class ExcelSignatureLineControllerTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAZAAAACWCAIAAADwkd5lAAAABmJLR0QA/wD/AP+gvaeTAAAACXBIWXMAABYlAAAWJQFJUiTwAAAB60lEQVR42u3TMQ0AAAgDIN8/9K3hHFQg3DQakJzZAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwDkGxAAB27JxWQAAAABJRU5ErkJggg==");

  @Test
  void controllerHandlesEmptySheetsLifecycleAndPreviewReferenceTracking() throws Exception {
    ExcelSignatureLineController controller = new ExcelSignatureLineController();
    ExcelDrawingAnchor.TwoCell initialAnchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);
    ExcelDrawingAnchor.TwoCell movedAnchor = ExcelChartTestSupport.anchor(6, 2, 10, 7);

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");

      assertEquals(List.of(), controller.signatureLines(sheet));
      assertFalse(controller.hasNamedSignatureLine(sheet, "Missing"));
      assertFalse(controller.updateAnchorIfPresent(sheet, "Missing", initialAnchor));
      assertFalse(controller.deleteIfPresent(sheet, "Missing"));
      assertFalse(
          ExcelSignatureLineController.usesImagePart(
              sheet, PackagingURIHelper.createPartName("/xl/media/missing.png")));

      controller.setSignatureLine(sheet, definition("OpsSignature", initialAnchor));
      assertTrue(controller.hasNamedSignatureLine(sheet, "OpsSignature"));
      assertEquals("OpsSignature", controller.signatureLines(sheet).getFirst().name());

      XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
      CTShape shape = signatureShapes(vmlDrawing).getFirst();
      PackagePart previewPart =
          ExcelDrawingBinarySupport.relatedInternalPart(
              vmlDrawing.getPackagePart(), shape.getImagedataArray(0).getRelid());
      PackagePartName previewPartName = previewPart.getPartName();
      assertTrue(ExcelSignatureLineController.usesImagePart(sheet, previewPartName));
      assertFalse(
          ExcelSignatureLineController.usesImagePart(
              sheet, PackagingURIHelper.createPartName("/xl/media/other-preview.png")));

      assertTrue(controller.updateAnchorIfPresent(sheet, "OpsSignature", movedAnchor));
      assertEquals(movedAnchor, controller.signatureLines(sheet).getFirst().anchor());

      assertTrue(controller.deleteIfPresent(sheet, "OpsSignature"));
      assertFalse(controller.hasNamedSignatureLine(sheet, "OpsSignature"));
      assertFalse(controller.deleteIfPresent(sheet, "OpsSignature"));
      assertFalse(ExcelSignatureLineController.usesImagePart(sheet, previewPartName));
    }
  }

  @Test
  void controllerRejectsMissingClientDataAndMissingAnchors() throws Exception {
    ExcelSignatureLineController controller = new ExcelSignatureLineController();
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setSignatureLine(sheet, definition("OpsSignature", anchor));
      CTShape shape = signatureShapes(sheet.getVMLDrawing(false)).getFirst();
      shape.removeClientData(0);

      IllegalStateException readFailure =
          assertThrows(IllegalStateException.class, () -> controller.signatureLines(sheet));
      assertEquals(
          "Signature line 'OpsSignature' is missing VML clientData", readFailure.getMessage());

      IllegalStateException moveFailure =
          assertThrows(
              IllegalStateException.class,
              () -> controller.updateAnchorIfPresent(sheet, "OpsSignature", anchor));
      assertEquals(
          "Signature line 'OpsSignature' on sheet 'Ops' is missing VML clientData",
          moveFailure.getMessage());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setSignatureLine(sheet, definition("OpsSignature", anchor));
      CTClientData clientData =
          signatureShapes(sheet.getVMLDrawing(false)).getFirst().getClientDataArray(0);
      clientData.removeAnchor(0);

      IllegalStateException readFailure =
          assertThrows(IllegalStateException.class, () -> controller.signatureLines(sheet));
      assertEquals(
          "Signature line 'OpsSignature' is missing its VML anchor", readFailure.getMessage());

      IllegalStateException moveFailure =
          assertThrows(
              IllegalStateException.class,
              () -> controller.updateAnchorIfPresent(sheet, "OpsSignature", anchor));
      assertEquals(
          "Signature line 'OpsSignature' on sheet 'Ops' is missing its VML anchor",
          moveFailure.getMessage());
    }
  }

  @Test
  void controllerRejectsInvalidAnchorsAndDuplicateResolvedNames() throws Exception {
    ExcelSignatureLineController controller = new ExcelSignatureLineController();
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setSignatureLine(sheet, definition("OpsSignature", anchor));
      CTClientData clientData =
          signatureShapes(sheet.getVMLDrawing(false)).getFirst().getClientDataArray(0);
      clientData.setAnchorArray(0, "1, 2, 3");
      IllegalStateException malformedTokens =
          assertThrows(IllegalStateException.class, () -> controller.signatureLines(sheet));
      assertTrue(malformedTokens.getMessage().contains("invalid VML anchor"));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setSignatureLine(sheet, definition("OpsSignature", anchor));
      CTClientData clientData =
          signatureShapes(sheet.getVMLDrawing(false)).getFirst().getClientDataArray(0);
      clientData.setAnchorArray(0, "a, b, c, d, e, f, g, h");
      IllegalStateException malformedNumbers =
          assertThrows(IllegalStateException.class, () -> controller.signatureLines(sheet));
      assertTrue(malformedNumbers.getMessage().contains("invalid VML anchor"));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setSignatureLine(sheet, definition("First", anchor));
      controller.setSignatureLine(
          sheet, definition("Second", ExcelChartTestSupport.anchor(5, 1, 8, 6)));
      List<CTShape> shapes = signatureShapes(sheet.getVMLDrawing(false));
      shapes.getFirst().getImagedataArray(0).setTitle("Dup");
      shapes.getFirst().setAlt("Dup");
      shapes.get(1).getImagedataArray(0).setTitle("Dup");
      shapes.get(1).setAlt("Dup");

      IllegalArgumentException duplicateNames =
          assertThrows(
              IllegalArgumentException.class, () -> controller.hasNamedSignatureLine(sheet, "Dup"));
      assertEquals(
          "Multiple signature lines named 'Dup' exist on sheet 'Ops'", duplicateNames.getMessage());
    }
  }

  @Test
  void controllerFallsBackToSetupIdsAndSyntheticIndexesWhenNamesAreUnavailable() throws Exception {
    ExcelSignatureLineController controller = new ExcelSignatureLineController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      controller.setSignatureLine(
          sheet, definition("Original", ExcelChartTestSupport.anchor(1, 1, 4, 6)));
      CTShape shape = signatureShapes(sheet.getVMLDrawing(false)).getFirst();
      shape.getImagedataArray(0).setTitle(" ");
      shape.setAlt("Microsoft Office Signature Line...");
      shape.getSignaturelineArray(0).setId("Setup42");
      assertEquals("SignatureLine-Setup42", controller.signatureLines(sheet).getFirst().name());

      shape.unsetAlt();
      shape.getSignaturelineArray(0).setId("Setup99");
      assertEquals("SignatureLine-Setup99", controller.signatureLines(sheet).getFirst().name());

      shape.getSignaturelineArray(0).setId("");
      assertEquals("SignatureLine-1", controller.signatureLines(sheet).getFirst().name());
    }
  }

  @Test
  void controllerCoversLeanSignatureMetadataWithoutPreviewImage() throws Exception {
    ExcelSignatureLineController controller = new ExcelSignatureLineController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      ExcelSheetAnnotationSupport annotationSupport =
          new ExcelSheetAnnotationSupport(sheet, new ExcelDrawingController());
      annotationSupport.setComment("C3", new ExcelComment("Note", "GridGrind", false));

      controller.setSignatureLine(
          sheet,
          new ExcelSignatureLineDefinition(
              "LeanSignature",
              ExcelChartTestSupport.anchor(1, 1, 4, 6),
              true,
              "Follow the packet",
              "Ada Lovelace",
              null,
              null,
              null,
              null,
              null,
              null));

      CTShape shape = signatureShapes(sheet.getVMLDrawing(false)).getFirst();
      shape.removeImagedata(0);
      ExcelSignatureLineController.applyImageTitle(shape, "LeanSignature");
      shape.getSignaturelineArray(0).unsetAllowcomments();

      ExcelSignatureLineSnapshot snapshot = controller.signatureLines(sheet).getFirst();
      assertNull(snapshot.allowComments());
      assertEquals("LeanSignature", snapshot.name());
      assertTrue(controller.hasNamedSignatureLine(sheet, "LeanSignature"));
      assertFalse(
          ExcelSignatureLineController.usesImagePart(
              sheet, PackagingURIHelper.createPartName("/xl/media/missing.png")));
    }
  }

  private static ExcelSignatureLineDefinition definition(
      String name, ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelSignatureLineDefinition(
        name,
        anchor,
        false,
        "Review before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        null,
        "invalid",
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(PNG_PIXEL_BYTES));
  }

  private static List<CTShape> signatureShapes(XSSFVMLDrawing vmlDrawing) {
    List<CTShape> shapes = new ArrayList<>();
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        if (cursor.getObject() instanceof CTShape shape && shape.sizeOfSignaturelineArray() > 0) {
          shapes.add(shape);
        }
      }
    }
    return List.copyOf(shapes);
  }
}
