package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/**
 * DSC header/trailer and decimal-span controls require no PostScript interpreter or numeric arrays.
 */
class EpsStructureTest {
  private static final String HEADER = "%!PS-Adobe-3.0 EPSF-3.0\n";

  @Test
  void boundingBoxMustBeInTheHeaderOrAnExplicitAtendTrailer() {
    assertDoesNotThrow(() ->
        admit(
            HEADER
                + "%%BoundingBox: (atend)\n"
                + "%%EndComments\n"
                + "0 pop\n"
                + "%%Trailer\n"
                + "%%BoundingBox: 0 0 0 0\n"));
    assertThrows(
        InvalidRequestException.class, () -> admit(HEADER + "0 pop\n%%BoundingBox: 0 0 0 0\n"));
    assertThrows(
        InvalidRequestException.class,
        () -> admit(HEADER + "%%BoundingBox: (atend)\n0 pop\n%%BoundingBox: 0 0 0 0\n"));
    assertThrows(
        InvalidRequestException.class,
        () -> admit(HEADER + "%%BoundingBox: (atend)\n(%%Trailer) pop\n%%BoundingBox: 0 0 0 0\n"));
  }

  @Test
  void signedVectorCoordinatesDoNotInheritAnInterpreterOrRasterIntegerCap() {
    assertDoesNotThrow(() ->
        admit(
            HEADER
                + "%%BoundingBox: -2147483648 -1 2147483648 1\n"
                + "newpath -2147483648 0 moveto 2147483648 0 lineto stroke\n"));
    assertDoesNotThrow(() -> admit("%!PS-Adobe-2.0\tEPSF-1.2\n%%BoundingBox: -0 +000 0 0\n"));
    assertThrows(
        InvalidRequestException.class, () -> admit(HEADER + "%%BoundingBox: -100 0 -200 1\n"));
  }

  @Test
  void hugeCountTokensFailTheActualResourceOrSpanBoundBeforeProportionalWork() {
    var pixels =
        assertThrows(
            InvalidRequestException.class,
            () ->
                admit(
                    HEADER
                        + "%%BoundingBox: 0 0 0 0\n"
                        + "%%BeginPreview: 999999999999999999999999999 1 1 1\n"
                        + "%00\n"
                        + "%%EndPreview\n"));
    var field = (FieldValidationProblem) pixels.requestProblem();
    assertEquals(FieldValidationMediaRule.RESOURCE_LIMIT, field.rule());
    assertEquals("LIM-045", field.operands().getFirst());
    assertThrows(
        InvalidRequestException.class,
        () ->
            admit(
                HEADER
                    + "%%BoundingBox: 0 0 0 0\n"
                    + "%%BeginBinary: 999999999999999999999999999\n"
                    + "%%EndBinary\n"));
  }

  @Test
  void malformedVersionClaimsAndRepeatedHeaderBoxesRefuse() {
    assertThrows(
        InvalidRequestException.class,
        () -> admit("%!PS-Adobe-foo EPSF-3.0\n%%BoundingBox: 0 0 0 0\n"));
    assertThrows(
        InvalidRequestException.class,
        () -> admit(HEADER + "%%BoundingBox: 0 0 0 0\n%%BoundingBox: 0 0 1 1\n"));
    assertThrows(
        InvalidRequestException.class,
        () -> admit(HEADER + "%%BoundingBox: 0 0 0 0\n%%BeginBinary: 0\n%%EndBinaryX\n"));
  }

  private static void admit(String program) {
    AuthoredImageAdmission.admitNative(
        ExcelPictureFormat.EPS, new BinaryContent(program.getBytes(StandardCharsets.US_ASCII)));
  }
}
