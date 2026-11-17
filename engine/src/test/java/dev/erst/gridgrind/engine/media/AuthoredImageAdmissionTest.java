package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Independent format controls and malformed-source boundaries exercise actual media ownership. */
class AuthoredImageAdmissionTest {
  @Test
  void everyDeclaredFormatRefusesTheOriginalNonimagePayload() {
    var content = new BinaryContent("This is not an image.\n".getBytes(StandardCharsets.US_ASCII));
    for (var format : ExcelPictureFormat.values()) {
      var failure =
          assertThrows(
              InvalidRequestException.class,
              () -> AuthoredImageAdmission.admitNative(format, content),
              format.name());
      assertEquals(
          FieldValidationMediaRule.FORMAT_STRUCTURE,
          ((FieldValidationProblem) failure.requestProblem()).rule());
    }
  }

  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void elevenFormatsPreserveExactPrivatelyAdmittedBytes() throws Exception {
    for (var format : ExcelPictureFormat.values()) {
      String extension =
          switch (format) {
            case JPEG -> "jpeg";
            case TIFF -> "tiff";
            default -> format.name().toLowerCase(java.util.Locale.ROOT);
          };
      byte[] original = fixture("valid." + extension);
      var source = new BinaryContent(original);
      var admitted = AuthoredImageAdmission.admitNative(format, source);
      assertArrayEquals(original, admitted.nativeBytes(format, source));
      byte[] nativeCopy = admitted.nativeBytes(format, source);
      nativeCopy[0] ^= 1;
      assertArrayEquals(original, admitted.nativeBytes(format, source));
      assertThrows(
          UnexpectedNativeContentException.class,
          () -> admitted.nativeBytes(format, new BinaryContent(original)));
    }
  }

  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void validCodingAndContainerVariantsRemainAdmitted() throws Exception {
    for (String name :
        List.of(
            "valid-adam7.png",
            "valid-animated.gif",
            "valid-progressive.jpeg",
            "valid-standard.wmf",
            "valid-v1.pict",
            "valid-v2.wpg",
            "valid-bigtiff.tiff",
            "valid-bigendian.tiff",
            "deflate.tiff",
            "deflate-legacy-tag.tiff",
            "packbits.tiff",
            "lzw-full-dictionary.tiff",
            "lzw-widths.tiff",
            "lzw-planar.tiff",
            "lzw-tiled.tiff",
            "lzw-bigendian.tiff",
            "g3-1d-msb2lsb-31-13.tif",
            "g3-2d-lsb2msb-31-13.tif",
            "g3-2d-fill-msb2lsb-31-13.tif",
            "g4-lsb2msb-31-13.tif",
            "old-jpeg-interchange.tif")) {
      ExcelPictureFormat format =
          switch (name.substring(name.lastIndexOf('.') + 1)) {
            case "png" -> ExcelPictureFormat.PNG;
            case "gif" -> ExcelPictureFormat.GIF;
            case "jpeg" -> ExcelPictureFormat.JPEG;
            case "wmf" -> ExcelPictureFormat.WMF;
            case "pict" -> ExcelPictureFormat.PICT;
            case "wpg" -> ExcelPictureFormat.WPG;
            default -> ExcelPictureFormat.TIFF;
          };
      var source = new BinaryContent(fixture(name));
      assertDoesNotThrow(() -> AuthoredImageAdmission.admitNative(format, source), name);
    }
  }

  @Test
  void changedDeclaredCodecMismatchAndTruncationRefuseForTheirStructure() throws Exception {
    var lzw =
        assertThrows(
            InvalidRequestException.class,
            () ->
                AuthoredImageAdmission.admitNative(
                    ExcelPictureFormat.TIFF,
                    new BinaryContent(fixture("framed-invalid-lzw.tiff"))));
    var field = (FieldValidationProblem) lzw.requestProblem();
    assertEquals(FieldValidationMediaRule.FORMAT_STRUCTURE, field.rule());
    assertEquals("SEQUENCE", field.operands().get(1));
    var png = new BinaryContent(fixture("valid.png"));
    assertThrows(
        InvalidRequestException.class,
        () -> AuthoredImageAdmission.admitNative(ExcelPictureFormat.JPEG, png));
    assertThrows(
        InvalidRequestException.class,
        () ->
            AuthoredImageAdmission.admitNative(
                ExcelPictureFormat.PNG, new BinaryContent(Arrays.copyOf(png.bytes(), 20))));
  }

  @Test
  void nativeUnexpectedArgumentFailuresAndFatalIdentityRemainVisible() {
    var original = new IllegalArgumentException("synthetic native failure");
    var wrapped =
        assertThrows(
            UnexpectedNativeContentException.class,
            () ->
                AuthoredImageAdmission.mutate(() -> {
                  throw original;
                }));
    assertSame(original, wrapped.getCause());
    var context =
        new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.unknown());
    assertEquals(
        dev.erst.gridgrind.core.problem.GridGrindProblemCode.INTERNAL_ERROR,
        dev.erst.gridgrind.engine.api.GridGrindProblems.fromException(wrapped, context).code());
    var fatal = new AssertionError("synthetic fatal failure");
    assertSame(
        fatal,
        assertThrows(
            AssertionError.class,
            () ->
                AuthoredImageAdmission.mutate(() -> {
                  throw fatal;
                })));
  }

  static byte[] fixture(String name) throws IOException {
    try (var input =
        java.util.Objects.requireNonNull(
            AuthoredImageAdmissionTest.class.getResourceAsStream("/media/" + name))) {
      return input.readAllBytes();
    }
  }
}
