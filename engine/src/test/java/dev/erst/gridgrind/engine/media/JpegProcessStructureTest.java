package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.List;
import org.junit.jupiter.api.Test;

/**
 * Genuine JPEG.org reference streams cover rare declared processes and their actual scan grammar.
 */
class JpegProcessStructureTest {
  @Test
  void referenceHierarchicalDifferentialDnlAndRestartControlsHaveTheirOriginalDimensions()
      throws Exception {
    for (String name :
        List.of(
            "hierarchical-arithmetic",
            "hierarchical-progressive-arithmetic",
            "hierarchical-hybrid-arithmetic",
            "hierarchical-sequential",
            "hierarchical-progressive",
            "hierarchical-hybrid",
            "dnl",
            "restart")) {
      var geometry = read("reference-" + name + ".jpeg");
      assertEquals(17, geometry.width(), name);
      assertEquals(19, geometry.height(), name);
    }
  }

  @Test
  void arithmeticLosslessMayDiscardEveryFinalZeroByte() throws Exception {
    var geometry = read("reference-lossless-arithmetic.jpeg");
    assertEquals(1, geometry.width());
    assertEquals(1, geometry.height());
    assertEquals(8, geometry.sampleBits());
  }

  @Test
  void differentialLosslessRequiresPredictorZeroAndHuffmanStillRequiresEntropyBits() {
    for (String file :
        List.of(
            "jpeg-hierarchical-hybrid-wrong-predictor.jpeg",
            "jpeg-hierarchical-hybrid-arithmetic-wrong-predictor.jpeg",
            "jpeg-huffman-empty-entropy.jpeg"))
      assertThrows(InvalidRequestException.class, () -> read(file), file);
  }

  @Test
  void interleavedSamplingBoundaryAndNoninterleavedHighSamplingRemainAdmitted() throws Exception {
    for (String file :
        List.of("jpeg-header-interleaved-ten.jpeg", "jpeg-header-noninterleaved-sixteen.jpeg")) {
      var geometry = read(file);
      assertEquals(17, geometry.width(), file);
      assertEquals(19, geometry.height(), file);
    }
  }

  @Test
  void scanOrderFollowsComponentDeclarationRatherThanNumericIdentifiers() throws Exception {
    var geometry = read("jpeg-header-declaration-order.jpeg");
    assertEquals(3, geometry.width());
    assertEquals(2, geometry.height());
  }

  @Test
  void malformedHeaderRelationshipsRefuseWhileTheirOriginalStreamsRemainAdmitted()
      throws Exception {
    assertEquals(3, read("valid.jpeg").width());
    assertEquals(17, read("reference-hierarchical-sequential.jpeg").width());
    for (String file :
        List.of(
            "jpeg-header-interleaved-overflow.jpeg",
            "jpeg-header-scan-order.jpeg",
            "jpeg-header-dhp-quant.jpeg",
            "jpeg-header-dhp-precision.jpeg")) {
      assertThrows(InvalidRequestException.class, () -> read(file), file);
    }
  }

  private static ImageGeometry read(String file) throws java.io.IOException {
    byte[] bytes = AuthoredImageAdmissionTest.fixture(file);
    var source = new BinaryContent(bytes);
    AuthoredImageAdmission.admitNative(ExcelPictureFormat.JPEG, source);
    return ImageStructures.check(
        new MediaInput(
            bytes, 0, bytes.length, ExcelPictureFormat.JPEG, new MediaAdmissionBudget(), 0));
  }
}
