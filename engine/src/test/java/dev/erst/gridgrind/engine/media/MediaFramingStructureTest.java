package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/**
 * A grammar floor accepts legal blank/control-only streams while refusing genuinely missing data.
 */
class MediaFramingStructureTest {
  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void everyFormatRequiresItsOwnNecessaryStructureBeyondAMagicPrefix() throws Exception {
    for (var format : ExcelPictureFormat.values()) {
      String suffix =
          switch (format) {
            case JPEG -> "jpeg";
            case TIFF -> "tiff";
            default -> format.name().toLowerCase(java.util.Locale.ROOT);
          };
      byte[] valid = AuthoredImageAdmissionTest.fixture("valid." + suffix);
      byte[] incomplete;
      if (format == ExcelPictureFormat.EPS) {
        String source = new String(valid, StandardCharsets.US_ASCII);
        incomplete =
            source.replace("%%BoundingBox: 0 0 3 2\n", "").getBytes(StandardCharsets.US_ASCII);
        assertFalse(Arrays.equals(valid, incomplete));
      } else incomplete = Arrays.copyOf(valid, valid.length / 2);
      var failure =
          assertThrows(
              InvalidRequestException.class,
              () -> AuthoredImageAdmission.admitNative(format, new BinaryContent(incomplete)),
              format.name());
      assertEquals(
          FieldValidationMediaRule.FORMAT_STRUCTURE,
          ((FieldValidationProblem) failure.requestProblem()).rule());
    }
  }

  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void standardBlankAndControlOnlyVariantsDoNotNeedRasterDecodingOrAnInventedEof()
      throws Exception {
    for (String file :
        List.of("blank.eps", "control-only.gif", "blank.wmf", "blank-v1.pict", "blank-v2.pict")) {
      ExcelPictureFormat format =
          file.endsWith("eps")
              ? ExcelPictureFormat.EPS
              : file.endsWith("gif")
                  ? ExcelPictureFormat.GIF
                  : file.endsWith("wmf") ? ExcelPictureFormat.WMF : ExcelPictureFormat.PICT;
      var source = new BinaryContent(AuthoredImageAdmissionTest.fixture(file));
      assertDoesNotThrow(() -> AuthoredImageAdmission.admitNative(format, source), file);
    }
  }
}
