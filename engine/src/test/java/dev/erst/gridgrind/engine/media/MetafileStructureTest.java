package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import org.junit.jupiter.api.Test;

/** Bounded opaque description extents do not assert an unqualified counted-NUL convention. */
class MetafileStructureTest {
  @Test
  void countedDescriptionWithFollowingHeaderNulPaddingRemainsByteExact() throws Exception {
    byte[] bytes = AuthoredImageAdmissionTest.fixture("counted-description.emf");
    var source = new BinaryContent(bytes);
    var admitted = AuthoredImageAdmission.admitNative(ExcelPictureFormat.EMF, source);
    assertArrayEquals(bytes, admitted.nativeBytes(ExcelPictureFormat.EMF, source));
  }

  @Test
  void descriptionCountAndOffsetRemainBoundedByTheHeaderRecord() throws Exception {
    byte[] count = AuthoredImageAdmissionTest.fixture("counted-description.emf");
    for (int i = 60; i < 64; i++) count[i] = (byte) 255;
    var tooMany = assertThrows(InvalidRequestException.class, () -> admit(count));
    assertEquals("OFFSET", ((FieldValidationProblem) tooMany.requestProblem()).operands().get(1));
    byte[] offset = AuthoredImageAdmissionTest.fixture("counted-description.emf");
    offset[64] = 120;
    var outside = assertThrows(InvalidRequestException.class, () -> admit(offset));
    assertEquals("OFFSET", ((FieldValidationProblem) outside.requestProblem()).operands().get(1));
  }

  private static void admit(byte[] bytes) {
    AuthoredImageAdmission.admitNative(ExcelPictureFormat.EMF, new BinaryContent(bytes));
  }
}
