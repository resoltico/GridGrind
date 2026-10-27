package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.nio.file.Path;
import java.util.Map;
import java.util.Set;
import javax.imageio.ImageIO;
import org.junit.jupiter.api.Test;
import org.opentest4j.AssertionFailedError;
import tools.jackson.databind.node.ObjectNode;

/** Negative controls for portable preview fingerprints and independent persisted-image checks. */
class GoldenSignaturePreviewsTest {
  @Test
  void persistedBytesDimensionsAndFingerprintMustAllAgree() throws Exception {
    byte[] image = png(0xFFFFFF);
    String hash = GoldenSignaturePreviews.sha256(image);
    ObjectNode preview = preview(image);
    GoldenSignaturePreviews.verifyPreview(preview, image, Map.of(hash, image));
    assertThrows(
        AssertionFailedError.class,
        () ->
            GoldenSignaturePreviews.verifyPreview(
                preview.deepCopy().put("byteSize", image.length + 1), image, Map.of(hash, image)));
    assertThrows(
        AssertionFailedError.class,
        () ->
            GoldenSignaturePreviews.verifyPreview(
                preview.deepCopy().put("widthPixels", 3), image, Map.of(hash, image)));
    assertThrows(
        AssertionFailedError.class,
        () ->
            GoldenSignaturePreviews.verifyPreview(
                preview.deepCopy().put("sha256", "0".repeat(64)), image, Map.of(hash, image)));
    byte[] differentPixels = png(0x0000FF);
    assertThrows(
        AssertionFailedError.class,
        () ->
            GoldenSignaturePreviews.verifyPreview(
                preview(differentPixels),
                image,
                Map.of(GoldenSignaturePreviews.sha256(differentPixels), differentPixels)));
  }

  @Test
  void normalizationKeepsPreviewIdentityAcrossExecutionAndReadback() throws Exception {
    ObjectNode baseline = observation("a".repeat(64), 100);
    ObjectNode differentRasterization = observation("b".repeat(64), 200);
    Path workspace = Path.of(".").toAbsolutePath();
    assertEquals(
        GoldenSnapshots.canonical(baseline, workspace, Set.of("a".repeat(64))),
        GoldenSnapshots.canonical(differentRasterization, workspace, Set.of("b".repeat(64))));
    assertNotEquals(
        GoldenSnapshots.canonical(baseline, workspace),
        GoldenSnapshots.canonical(differentRasterization, workspace),
        "Uncertified fixed asset previews must keep their exact fingerprints");
    ObjectNode inconsistent = baseline.deepCopy();
    ((ObjectNode) inconsistent.path("after").path("preview")).put("sha256", "c".repeat(64));
    assertNotEquals(
        GoldenSnapshots.canonical(baseline, workspace, Set.of("a".repeat(64))),
        GoldenSnapshots.canonical(inconsistent, workspace, Set.of("a".repeat(64))));
    ObjectNode wrongSize = baseline.deepCopy();
    ((ObjectNode) wrongSize.path("after").path("preview")).put("byteSize", 101);
    assertThrows(
        AssertionFailedError.class,
        () -> GoldenSnapshots.canonical(wrongSize, workspace, Set.of("a".repeat(64))));
    ObjectNode wrongType = baseline.deepCopy();
    ((ObjectNode) wrongType.path("after").path("preview")).put("byteSize", "100");
    assertThrows(
        AssertionFailedError.class,
        () -> GoldenSnapshots.canonical(wrongType, workspace, Set.of("a".repeat(64))));
  }

  private static ObjectNode observation(String hash, int size) {
    ObjectNode observation = GoldenSnapshots.JSON.createObjectNode();
    for (String phase : new String[] {"before", "after"}) {
      ObjectNode drawing = observation.putObject(phase).put("type", "SIGNATURE_LINE");
      drawing
          .putObject("preview")
          .put("sha256", hash)
          .put("byteSize", size)
          .put("format", "PNG")
          .put("widthPixels", 2)
          .put("heightPixels", 3);
    }
    return observation;
  }

  private static ObjectNode preview(byte[] image) {
    return GoldenSnapshots.JSON
        .createObjectNode()
        .put("format", "PNG")
        .put("contentType", "image/png")
        .put("sha256", GoldenSignaturePreviews.sha256(image))
        .put("byteSize", image.length)
        .put("widthPixels", 2)
        .put("heightPixels", 3);
  }

  private static byte[] png(int color) throws Exception {
    var image = new BufferedImage(2, 3, BufferedImage.TYPE_INT_RGB);
    image.setRGB(0, 0, color);
    var bytes = new ByteArrayOutputStream();
    assertTrue(ImageIO.write(image, "PNG", bytes));
    return bytes.toByteArray();
  }
}
