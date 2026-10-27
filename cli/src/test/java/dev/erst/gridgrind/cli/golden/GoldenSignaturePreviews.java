package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.Base64;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.zip.ZipFile;
import javax.imageio.ImageIO;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSignatureLine;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tools.jackson.databind.JsonNode;

/**
 * Checks uniquely named signature fixtures against native rendering in the current font
 * environment.
 */
final class GoldenSignaturePreviews {
  private GoldenSignaturePreviews() {}

  static Set<String> certify(Path request, Path workspace, JsonNode execution) throws IOException {
    JsonNode plan = GoldenSnapshots.JSON.readTree(Files.readAllBytes(request));
    JsonNode persistence = plan.path("persistence");
    if ("NONE".equals(persistence.path("type").asString())) return Set.of();
    String path =
        "SAVE_AS".equals(persistence.path("type").asString())
            ? persistence.path("path").asString()
            : plan.path("source").path("path").asString();
    return verify(plan, workspace, workspace.resolve(path), execution.path("readback"));
  }

  static Set<String> verify(JsonNode plan, Path workspace, Path output, JsonNode readback)
      throws IOException {
    Map<String, byte[]> expected = new LinkedHashMap<>();
    for (JsonNode step : plan.path("steps")) {
      JsonNode action = step.path("action");
      if ("SET_SIGNATURE_LINE".equals(action.path("type").asString(""))) {
        JsonNode definition = action.path("signatureLine");
        String name = definition.path("name").asString();
        assertTrue(
            expected.putIfAbsent(name, render(definition, workspace)) == null,
            "Golden signature fixtures require unique names: " + name);
      }
    }
    if (expected.isEmpty()) return Set.of();
    Map<String, byte[]> media = new LinkedHashMap<>();
    try (var zip = new ZipFile(output.toFile())) {
      for (var entry :
          zip.stream()
              .filter(e -> !e.isDirectory() && e.getName().startsWith("xl/media/"))
              .toList()) {
        try (var input = zip.getInputStream(entry)) {
          byte[] bytes = input.readAllBytes();
          media.put(sha256(bytes), bytes);
        }
      }
    }
    Set<String> certified = new LinkedHashSet<>();
    verifyReports(readback, expected, media, certified);
    return Set.copyOf(certified);
  }

  private static byte[] render(JsonNode definition, Path workspace) throws IOException {
    var signature = new XSSFSignatureLine();
    if (definition.has("suggestedSigner"))
      signature.setSuggestedSigner(definition.path("suggestedSigner").asString());
    if (definition.has("suggestedSigner2"))
      signature.setSuggestedSigner2(definition.path("suggestedSigner2").asString());
    if (definition.has("suggestedSignerEmail"))
      signature.setSuggestedSignerEmail(definition.path("suggestedSignerEmail").asString());
    if (definition.has("signingInstructions"))
      signature.setSigningInstructions(definition.path("signingInstructions").asString());
    if (definition.has("caption")) signature.setCaption(definition.path("caption").asString());
    if (definition.has("invalidStamp"))
      signature.setInvalidStamp(definition.path("invalidStamp").asString());
    if (definition.has("plainSignature")) {
      JsonNode picture = definition.path("plainSignature");
      assertEquals(
          "PNG", picture.path("format").asString(), "Golden signature fixture image format");
      JsonNode source = picture.path("source");
      byte[] bytes =
          switch (source.path("type").asString()) {
            case "INLINE_BASE64" ->
                Base64.getDecoder().decode(source.path("base64Data").asString());
            case "FILE" -> Files.readAllBytes(workspace.resolve(source.path("path").asString()));
            default ->
                throw new IllegalArgumentException(
                    "Golden signature fixture needs an inline or file image");
          };
      signature.setPlainSignature(bytes);
      signature.setContentType("image/png");
    }
    try (var workbook = new XSSFWorkbook()) {
      signature.add(
          workbook.createSheet("Reference"), new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 3, 5));
      return workbook.getAllPictures().getFirst().getData();
    }
  }

  private static void verifyReports(
      JsonNode node, Map<String, byte[]> expected, Map<String, byte[]> media, Set<String> certified)
      throws IOException {
    if (node.isObject()) {
      if ("SIGNATURE_LINE".equals(node.path("type").asString("")) && node.has("preview")) {
        String name = node.path("name").asString();
        if (expected.containsKey(name)) {
          verifyPreview(node.path("preview"), expected.get(name), media);
          certified.add(node.path("preview").path("sha256").asString());
        }
      }
      for (var property : node.properties())
        verifyReports(property.getValue(), expected, media, certified);
    } else if (node.isArray()) {
      for (JsonNode value : node) verifyReports(value, expected, media, certified);
    }
  }

  static void verifyPreview(JsonNode preview, byte[] expected, Map<String, byte[]> media)
      throws IOException {
    assertEquals("PNG", preview.path("format").asString());
    assertEquals("image/png", preview.path("contentType").asString());
    assertTrue(preview.path("byteSize").isIntegralNumber(), "Preview byte size must be an integer");
    String hash = preview.path("sha256").asString();
    assertTrue(media.containsKey(hash), "Reported preview SHA-256 must identify persisted bytes");
    byte[] bytes = media.get(hash);
    assertArrayEquals(
        expected,
        bytes,
        "Persisted signature preview must match the native rendering of the authored fields");
    assertEquals(bytes.length, preview.path("byteSize").asLong());
    var image =
        Objects.requireNonNull(
            ImageIO.read(new ByteArrayInputStream(bytes)), "Preview must decode as an image");
    assertEquals(image.getWidth(), preview.path("widthPixels").asInt());
    assertEquals(image.getHeight(), preview.path("heightPixels").asInt());
  }

  static String sha256(byte[] bytes) {
    try {
      return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes));
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 is required by the JDK", exception);
    }
  }
}
