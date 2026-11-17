package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.drawing.SignatureAdapter;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/**
 * Actual POI mutation boundaries guard every authored image route before deleting existing objects.
 */
class AuthoredMediaNativeBoundaryTest {
  private static final DrawingAnchorSpec.TwoCell ANCHOR =
      DrawingAnchorSpec.TwoCell.moveAndResize(
          new DrawingMarkerSpec(0, 0), new DrawingMarkerSpec(2, 2));

  @Test
  void malformedPictureLeavesTheNamedShapeUntouched() throws Exception {
    assertShapeRemains((drawings, sheet) ->
        drawings.setPicture(
            sheet,
            new ResolvedPicture(
                "Keep", malformed(), ExcelPictureFormat.PNG, ANCHOR, Optional.empty())));
  }

  @Test
  void malformedEmbeddedPreviewLeavesTheNamedShapeAndPackageUntouched() throws Exception {
    assertShapeRemains((drawings, sheet) ->
        drawings.setEmbeddedObject(
            sheet,
            new ResolvedEmbeddedObject(
                "Keep",
                "Payload",
                "sample.bin",
                "open",
                new BinaryContent(new byte[] {4}),
                ExcelPictureFormat.PNG,
                malformed(),
                ANCHOR)));
  }

  @Test
  void malformedOuterSignatureLeavesTheNamedShapeUntouched() throws Exception {
    assertShapeRemains(
        (drawings, sheet) -> drawings.setSignatureLine(sheet, signature(malformed())));
  }

  private static void assertShapeRemains(
      java.util.function.BiConsumer<DrawingAdapter, org.apache.poi.xssf.usermodel.XSSFSheet>
          operation)
      throws java.io.IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
    try (var workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Sheet1");
      var drawings = new DrawingAdapter();
      drawings.setShape(
          sheet, new ResolvedShape.SimpleShape("Keep", ANCHOR, "rect", Optional.empty()));
      String original = sheet.getDrawingPatriarch().getCTDrawing().xmlText();
      int parts = workbook.getPackage().getParts().size();
      assertThrows(InvalidRequestException.class, () -> operation.accept(drawings, sheet));
      assertEquals(original, sheet.getDrawingPatriarch().getCTDrawing().xmlText());
      assertEquals(parts, workbook.getPackage().getParts().size());
      assertNull(sheet.getVMLDrawing(false));
    }
  }

  private static BinaryContent malformed() {
    return new BinaryContent(new byte[] {1, 2, 3});
  }

  @Test
  void publicInnerSignatureGuardPreservesExistingVmlAndPreviewParts() throws Exception {
    try (var workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Sheet1");
      var signatures = new SignatureAdapter();
      signatures.setSignatureLine(
          sheet, signature(new BinaryContent(AuthoredImageAdmissionTest.fixture("valid.png"))));
      String original = sheet.getVMLDrawing(false).getDocument().getXml().xmlText();
      int parts = workbook.getPackage().getParts().size();
      assertThrows(
          InvalidRequestException.class,
          () ->
              signatures.setSignatureLine(
                  sheet, signature(new BinaryContent(new byte[] {1, 2, 3}))));
      assertEquals(original, sheet.getVMLDrawing(false).getDocument().getXml().xmlText());
      assertEquals(parts, workbook.getPackage().getParts().size());
    }
  }

  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void pictureStorageKeepsAllElevenDeclaredFormatsAndExactPayloads() throws Exception {
    try (var workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Sheet1");
      var drawings = new DrawingAdapter();
      for (var format : ExcelPictureFormat.values()) {
        String suffix =
            switch (format) {
              case JPEG -> "jpeg";
              case TIFF -> "tiff";
              default -> format.name().toLowerCase(java.util.Locale.ROOT);
            };
        byte[] bytes = AuthoredImageAdmissionTest.fixture("valid." + suffix);
        drawings.setPicture(
            sheet,
            new ResolvedPicture(
                format.name(), new BinaryContent(bytes), format, ANCHOR, Optional.empty()));
        var picture = workbook.getAllPictures().getLast();
        assertArrayEquals(bytes, picture.getData());
        assertEquals(format.defaultContentType(), picture.getPackagePart().getContentType());
      }
      assertEquals(11, workbook.getAllPictures().size());
    }
  }

  private static ResolvedSignatureLine signature(BinaryContent preview) {
    return new ResolvedSignatureLine(
        "Keep",
        ANCHOR,
        false,
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.of("Synthetic signer"),
        Optional.empty(),
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(preview));
  }
}
