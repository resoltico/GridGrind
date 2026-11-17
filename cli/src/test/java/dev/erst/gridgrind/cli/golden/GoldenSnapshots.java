package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
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
    verify(id, launcher, workspace, observation, Set.of(), Set.of());
  }

  static void verify(
      String id,
      GoldenLauncher launcher,
      Path workspace,
      ObjectNode observation,
      Set<String> certifiedGeneratedPreviews,
      Set<String> certifiedGeneratedSetupIds)
      throws IOException {
    String actual =
        canonical(observation, workspace, certifiedGeneratedPreviews, certifiedGeneratedSetupIds);
    Path snapshot = DIRECTORY.resolve(id + ".json");
    if (launcher == GoldenLauncher.IN_PROCESS && Boolean.getBoolean("gridgrind.golden.update")) {
      assertEquals(
          "0.0.0-dev",
          System.getProperty("gridgrind.test.expectedVersion"),
          "Record the shared golden corpus with the development version");
      Files.createDirectories(snapshot.getParent());
      Files.writeString(snapshot, actual);
    }
    Path stderr = snapshot.resolveSibling(id.substring(id.lastIndexOf('/') + 1) + ".stderr.json");
    if (Files.isRegularFile(stderr)) {
      JsonNode expected =
          JSON.readTree(
              expectedForVersion(
                  Files.readString(snapshot),
                  System.getProperty("gridgrind.test.expectedVersion")));
      JsonNode byLauncher = JSON.readTree(Files.readString(stderr));
      assertEquals(
          expectedForLauncher(expected, byLauncher, launcher),
          JSON.readTree(actual),
          () -> "Golden launcher drift: " + snapshot + " / " + launcher);
    } else {
      assertMatches(snapshot, actual);
    }
  }

  static JsonNode expectedForLauncher(JsonNode common, JsonNode stderr, GoldenLauncher launcher) {
    assertEquals(
        Set.of("IN_PROCESS", "SHADOW_JAR"),
        stderr.properties().stream()
            .map(Map.Entry::getKey)
            .collect(java.util.stream.Collectors.toSet()),
        "Stderr expectations must explicitly cover exactly the known golden launchers");
    ObjectNode expected = (ObjectNode) common.deepCopy();
    ObjectNode primary =
        expected.has("execution") ? (ObjectNode) expected.path("execution") : expected;
    assertTrue(
        primary.has("stderr"),
        "Launcher-specific expectation needs an existing execution stderr field");
    primary.set("stderr", stderr.path(launcher.name()));
    return expected;
  }

  static void assertMatches(Path snapshot, String actual) throws IOException {
    assertTrue(Files.isRegularFile(snapshot), () -> "Missing golden snapshot: " + snapshot);
    assertEquals(
        expectedForVersion(
            Files.readString(snapshot), System.getProperty("gridgrind.test.expectedVersion")),
        actual,
        () -> "Golden drift: " + snapshot);
  }

  /**
   * Expand only expected build references; incorrect references in actual output remain visible.
   */
  static String expectedForVersion(String snapshot, String version) {
    if ("0.0.0-dev".equals(version)) {
      return snapshot;
    }
    return snapshot
        .replace(
            "https://github.com/resoltico/GridGrind/blob/main",
            "https://github.com/resoltico/GridGrind/blob/v" + version)
        .replace("ghcr.io/resoltico/gridgrind:latest", "ghcr.io/resoltico/gridgrind:" + version);
  }

  static ObjectNode observation(GoldenLauncher.Result result) {
    ObjectNode node = JSON.createObjectNode();
    node.put("exit", result.exit());
    node.set("stdout", payload(result.stdout()));
    node.set("stderr", payload(result.stderr()));
    return node;
  }

  static String canonical(JsonNode node, Path workspace) {
    return canonical(node, workspace, Set.of(), Set.of());
  }

  static String canonical(JsonNode node, Path workspace, Set<String> certifiedGeneratedPreviews) {
    return canonical(node, workspace, certifiedGeneratedPreviews, Set.of());
  }

  static String canonical(
      JsonNode node,
      Path workspace,
      Set<String> certifiedGeneratedPreviews,
      Set<String> certifiedGeneratedSetupIds) {
    return JSON.writerWithDefaultPrettyPrinter()
            .writeValueAsString(
                normalize(
                    node,
                    workspace,
                    new NormalizationState(certifiedGeneratedPreviews, certifiedGeneratedSetupIds)))
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
          .forEach(entry -> {
            String name = entry.getKey();
            JsonNode value = entry.getValue();
            if (VOLATILE_FIELDS.contains(name)
                || (node.has("executionPath") && Set.of("sha256", "byteSize").contains(name))) {
              normalized.put(name, "<" + name + ">");
            } else if ("setup".equals(name)
                && "SIGNATURE_LINE".equals(node.path("type").asString(""))) {
              normalized.set(name, signatureSetup(value, workspace, state));
            } else if ("preview".equals(name)
                && "SIGNATURE_LINE".equals(node.path("type").asString(""))
                && state.certifiedGeneratedPreviews.contains(value.path("sha256").asString(""))) {
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

  private static JsonNode signatureSetup(JsonNode value, Path workspace, NormalizationState state) {
    ObjectNode setup = (ObjectNode) normalize(value, workspace, state);
    String id = value.path("setupId").asString("");
    if (state.certifiedGeneratedSetupIds.contains(id)) {
      setup.put(
          "setupId",
          state.setupIds.computeIfAbsent(id, ignored -> "<setupId-" + state.setupIds.size() + ">"));
    }
    return setup;
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
  @SuppressWarnings(
      "PMD.UseConcurrentHashMap") // One canonicalization owns insertion-ordered identities.
  private static final class NormalizationState {
    private final Set<String> certifiedGeneratedPreviews;
    private final Set<String> certifiedGeneratedSetupIds;

    private NormalizationState(
        Set<String> certifiedGeneratedPreviews, Set<String> certifiedGeneratedSetupIds) {
      this.certifiedGeneratedPreviews = Set.copyOf(certifiedGeneratedPreviews);
      this.certifiedGeneratedSetupIds = Set.copyOf(certifiedGeneratedSetupIds);
    }

    private final Map<String, String> setupIds = new LinkedHashMap<>();
    private final Map<String, PreviewIdentity> previews = new LinkedHashMap<>();
  }

  /** One generated image identity with a consistent reported byte count. */
  private record PreviewIdentity(int index, long byteSize) {}
}
