package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.List;
import org.junit.jupiter.api.Test;

/** MS-EMF/WMF concrete DIB receiving contexts preserve selected scans and logical palette words. */
class MetafileBitmapStructureTest {
  @Test
  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  // Each bounded case requires independent immutable input/ownership or budget state.
  void realMetafilePathsRetainTheBitmapTransferAndPaletteContexts() throws Exception {
    for (var format : List.of(ExcelPictureFormat.EMF, ExcelPictureFormat.WMF))
      for (String member : List.of("partial-rgb", "logical-palette", "partial-logical-palette")) {
        String name = member + "." + format.name().toLowerCase(java.util.Locale.ROOT);
        byte[] bytes = AuthoredImageAdmissionTest.fixture(name);
        assertDoesNotThrow(
            () -> AuthoredImageAdmission.admitNative(format, new BinaryContent(bytes)), name);
      }
  }

  @Test
  void aPartialTransferStillRequiresEveryDeclaredPixelByte() {
    assertThrows(
        InvalidRequestException.class,
        () ->
            AuthoredImageAdmission.admitNative(
                ExcelPictureFormat.EMF,
                new BinaryContent(AuthoredImageAdmissionTest.fixture("short-partial-member.emf"))));
  }
}
