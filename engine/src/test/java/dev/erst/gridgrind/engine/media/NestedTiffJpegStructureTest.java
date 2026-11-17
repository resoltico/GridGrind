package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.List;
import org.junit.jupiter.api.Test;

/** TIFF-owned pointers keep their origin through shifted regions and actual DOS EPS previews. */
class NestedTiffJpegStructureTest {
  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded origin case needs independent byte storage and inspection counters.
  void identicalOldJpegTiffWorksAtTheTopLevelAndAtShiftedOrigins() throws Exception {
    byte[] tiff = AuthoredImageAdmissionTest.fixture("old-jpeg-interchange.tif");
    for (int shift : new int[] {0, 16, 74}) {
      byte[] bytes = new byte[tiff.length + shift];
      System.arraycopy(tiff, 0, bytes, shift, tiff.length);
      var geometry =
          TiffStructure.check(
              new MediaInput(
                  bytes,
                  shift,
                  tiff.length,
                  ExcelPictureFormat.TIFF,
                  new MediaAdmissionBudget(),
                  1));
      assertEquals(3, geometry.width());
      assertEquals(2, geometry.height());
    }
  }

  @Test
  void realDosEpsTiffPreviewPreservesTheExactPayload() throws Exception {
    byte[] bytes = AuthoredImageAdmissionTest.fixture("old-jpeg-preview.eps");
    var source = new BinaryContent(bytes);
    var admitted = AuthoredImageAdmission.admitNative(ExcelPictureFormat.EPS, source);
    assertArrayEquals(bytes, admitted.nativeBytes(ExcelPictureFormat.EPS, source));
  }

  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each pointer mutation is a distinct immutable payload and actual container admission.
  void entropyAndEveryTablePointerPlusOneRefuseInBothOwningContainers() throws Exception {
    for (String pointer : List.of("entropy", "quant", "dc", "ac")) {
      byte[] tiff = AuthoredImageAdmissionTest.fixture("old-jpeg-" + pointer + "-plus-one.tif");
      assertThrows(
          InvalidRequestException.class,
          () ->
              AuthoredImageAdmission.admitNative(ExcelPictureFormat.TIFF, new BinaryContent(tiff)),
          pointer);
      byte[] eps =
          AuthoredImageAdmissionTest.fixture("old-jpeg-" + pointer + "-plus-one-preview.eps");
      assertThrows(
          InvalidRequestException.class,
          () -> AuthoredImageAdmission.admitNative(ExcelPictureFormat.EPS, new BinaryContent(eps)),
          pointer);
    }
  }
}
