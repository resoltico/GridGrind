package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;
import tools.jackson.databind.node.StringNode;

/** Stores reviewed observations; masks only transport and artifact nondeterminism. */
final class GoldenSnapshots {
  static final JsonMapper JSON = JsonMapper.builder().build();
  private static final Path DIRECTORY =
      GoldenLauncher.REPOSITORY.resolve("cli/src/test/resources/golden/snapshots");
  private static final Set<String> VOLATILE_FIELDS =
      Set.of("timestamp", "startedAt", "finishedAt", "durationMillis", "version");

  private GoldenSnapshots() {}

  static void verify(String id, GoldenLauncher launcher, Path workspace, ObjectNode observation)
      throws IOException {
    verify(id, launcher, workspace, observation, Set.of());
  }

  static void verify(
      String id,
      GoldenLauncher launcher,
      Path workspace,
      ObjectNode observation,
      Set<String> certifiedGeneratedPreviews)
      throws IOException {
    String actual = canonical(observation, workspace, certifiedGeneratedPreviews);
    Path snapshot = DIRECTORY.resolve(id + ".json");
    if (launcher == GoldenLauncher.IN_PROCESS && Boolean.getBoolean("gridgrind.golden.update")) {
      Files.createDirectories(snapshot.getParent());
      Files.writeString(snapshot, actual);
    }
    assertMatches(snapshot, actual);
  }

  static void assertMatches(Path snapshot, String actual) throws IOException {
    assertTrue(Files.isRegularFile(snapshot), () -> "Missing golden snapshot: " + snapshot);
    assertEquals(Files.readString(snapshot), actual, () -> "Golden drift: " + snapshot);
  }

  static ObjectNode observation(GoldenLauncher.Result result) {
    ObjectNode node = JSON.createObjectNode();
    node.put("exit", result.exit());
    node.set("stdout", payload(result.stdout()));
    node.set("stderr", payload(result.stderr()));
    return node;
  }

  static String canonical(JsonNode node, Path workspace) {
    return canonical(node, workspace, Set.of());
  }

  static String canonical(JsonNode node, Path workspace, Set<String> certifiedGeneratedPreviews) {
    return JSON.writerWithDefaultPrettyPrinter()
            .writeValueAsString(
                normalize(node, workspace, new NormalizationState(certifiedGeneratedPreviews)))
        + "\n";
  }

  private static JsonNode payload(String text) {
    if (text.isBlank()) {
      return StringNode.valueOf(text);
    }
    try {
      return JSON.readTree(text);
    } catch (tools.jackson.core.JacksonException ignored) {
      ArrayNode lines = JSON.createArrayNode();
      for (String line : text.strip().lines().toList()) {
        try {
          lines.add(JSON.readTree(line));
        } catch (tools.jackson.core.JacksonException notJson) {
          return StringNode.valueOf(text);
        }
      }
      return lines;
    }
  }

  private static JsonNode normalize(JsonNode node, Path workspace, NormalizationState state) {
    if (node.isObject()) {
      ObjectNode normalized = JSON.createObjectNode();
      node.properties().stream()
          .sorted(Map.Entry.comparingByKey())
          .forEach(
              entry -> {
                String name = entry.getKey();
                JsonNode value = entry.getValue();
                if (VOLATILE_FIELDS.contains(name)
                    || (node.has("executionPath") && Set.of("sha256", "byteSize").contains(name))) {
                  normalized.put(name, "<" + name + ">");
                } else if ("setupId".equals(name) && value.isString()) {
                  String id = value.asString();
                  assertTrue(
                      id.startsWith("{") && id.endsWith("}"), "Signature setup id must be a UUID");
                  UUID.fromString(id.substring(1, id.length() - 1));
                  normalized.put(
                      name,
                      state.setupIds.computeIfAbsent(
                          id, ignored -> "<setupId-" + state.setupIds.size() + ">"));
                } else if ("preview".equals(name)
                    && "SIGNATURE_LINE".equals(node.path("type").asString(""))
                    && state.certifiedGeneratedPreviews.contains(
                        value.path("sha256").asString(""))) {
                  normalized.set(name, signaturePreview(value, workspace, state));
                } else {
                  normalized.set(name, normalize(value, workspace, state));
                }
              });
      return normalized;
    }
    if (node.isArray()) {
      ArrayNode normalized = JSON.createArrayNode();
      node.forEach(value -> normalized.add(normalize(value, workspace, state)));
      return normalized;
    }
    if (node.isString()) {
      String text = node.asString();
      String base = workspace.toAbsolutePath().normalize().toString();
      text = text.replace("/private" + base, "<workspace>").replace(base, "<workspace>");
      text = text.replace(GoldenLauncher.REPOSITORY.toString(), "<repository>");
      text = text.replaceAll("(?m)(GridGrind )\\d+\\.\\d+\\.\\d+(?:[-+][\\w.-]+)?", "$1<version>");
      return StringNode.valueOf(text);
    }
    return node;
  }

  private static JsonNode signaturePreview(
      JsonNode preview, Path workspace, NormalizationState state) {
    assertTrue(preview.path("sha256").isString(), "Signature preview hash must be a string");
    assertTrue(
        preview.path("byteSize").isIntegralNumber(),
        "Signature preview byte size must be an integer");
    String hash = preview.path("sha256").asString();
    long byteSize = preview.path("byteSize").asLong();
    assertTrue(hash.matches("[0-9a-f]{64}"), "Signature preview must carry a SHA-256");
    assertTrue(byteSize > 0, "Signature preview must contain image bytes");
    PreviewIdentity identity =
        state.previews.computeIfAbsent(
            hash, ignored -> new PreviewIdentity(state.previews.size(), byteSize));
    assertEquals(identity.byteSize(), byteSize, "The same preview hash must retain its byte size");
    ObjectNode normalized = (ObjectNode) normalize(preview, workspace, state);
    normalized.put("sha256", "<signature-preview-sha256-" + identity.index() + ">");
    normalized.put("byteSize", "<signature-preview-byteSize-" + identity.index() + ">");
    return normalized;
  }

  /** Identity relationships retained across one execution and its readback. */
  private static final class NormalizationState {
    private final Set<String> certifiedGeneratedPreviews;

    private NormalizationState(Set<String> certifiedGeneratedPreviews) {
      this.certifiedGeneratedPreviews = Set.copyOf(certifiedGeneratedPreviews);
    }

    private final Map<String, String> setupIds = new LinkedHashMap<>();
    private final Map<String, PreviewIdentity> previews = new LinkedHashMap<>();
  }

  /** One generated image identity with a consistent reported byte count. */
  private record PreviewIdentity(int index, long byteSize) {}
}
