package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.List;
import org.junit.jupiter.api.Test;

/**
 * Independent coding-profile controls retain required variants and reject actual malformed streams.
 */
class MediaProfileStructureTest {
  @Test
  void arithmeticLosslessAndTwelveBitJpegControlsHaveIndependentDimensions() throws Exception {
    for (String name :
        List.of(
            "arithmetic.jpeg",
            "progressive-arithmetic.jpeg",
            "lossless.jpeg",
            "lossless16.jpeg",
            "sequential12.jpeg")) {
      var geometry = inspect(name, ExcelPictureFormat.JPEG);
      assertEquals(3, geometry.width(), name);
      assertEquals(2, geometry.height(), name);
    }
  }

  @Test
  void os2HeaderSizesHotspotsHuffmanAndRle24FollowTheirOwnHeaderContract() throws Exception {
    for (String name : List.of("huffman-white8x2.bmp", "huffman-short-header.bmp")) {
      var geometry = inspect(name, ExcelPictureFormat.BMP);
      assertEquals(8, geometry.width());
      assertEquals(2, geometry.height());
    }
    var rle = inspect("rle24-literal5.bmp", ExcelPictureFormat.BMP);
    assertEquals(5, rle.width());
    assertEquals(1, rle.height());
    byte[] malformed = AuthoredImageAdmissionTest.fixture("huffman-white8x2.bmp");
    malformed[86] = 1;
    assertThrows(InvalidRequestException.class, () -> admit(ExcelPictureFormat.BMP, malformed));
  }

  @Test
  void subsampledYcbcrPackBitsSupportsStripsAndTilesWithoutFictitiousByteRows() throws Exception {
    for (String factor : List.of("1x1", "2x1", "2x2", "4x1", "4x2", "4x4"))
      for (String kind : List.of("strip", "tile")) {
        var geometry =
            inspect("ycbcr-packbits-" + factor + "-" + kind + ".tiff", ExcelPictureFormat.TIFF);
        assertEquals(5, geometry.width());
        assertEquals(7, geometry.height());
      }
    byte[] overrun = AuthoredImageAdmissionTest.fixture("ycbcr-packbits-2x2-strip.tiff");
    // The independent writer places the first strip at8. A repeat128 cannot fit its36decoded bytes.
    overrun[8] = (byte) 129;
    var failure =
        assertThrows(InvalidRequestException.class, () -> admit(ExcelPictureFormat.TIFF, overrun));
    assertEquals("COUNT", ((FieldValidationProblem) failure.requestProblem()).operands().get(1));
  }

  @Test
  void epsLexicalBoundsAndBalancedContainersAreAppliedWithoutExecution() {
    String header = "%!PS-Adobe-3.0 EPSF-3.0\n%%BoundingBox: 0 0 1 1\n";
    assertDoesNotThrow(() ->
        admit(
            ExcelPictureFormat.EPS,
            (header + "0 ".repeat(999_999)).getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
    var tokens =
        assertThrows(
            InvalidRequestException.class,
            () ->
                admit(
                    ExcelPictureFormat.EPS,
                    (header + "0 ".repeat(1_000_000))
                        .getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
    assertEquals(
        FieldValidationMediaRule.RESOURCE_LIMIT,
        ((FieldValidationProblem) tokens.requestProblem()).rule());
    assertDoesNotThrow(() ->
        admit(
            ExcelPictureFormat.EPS,
            (header + "[".repeat(256) + "]".repeat(256))
                .getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
    assertThrows(
        InvalidRequestException.class,
        () ->
            admit(
                ExcelPictureFormat.EPS,
                (header + "[".repeat(257) + "]".repeat(257))
                    .getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
    assertThrows(
        InvalidRequestException.class,
        () ->
            admit(
                ExcelPictureFormat.EPS,
                (header + "(unterminated").getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
  }

  private static void admit(ExcelPictureFormat format, byte[] bytes) {
    AuthoredImageAdmission.admitNative(format, new BinaryContent(bytes));
  }

  private static ImageGeometry inspect(String name, ExcelPictureFormat format)
      throws java.io.IOException {
    byte[] bytes = AuthoredImageAdmissionTest.fixture(name);
    return ImageStructures.check(
        new MediaInput(bytes, 0, bytes.length, format, new MediaAdmissionBudget(), 0));
  }
}
