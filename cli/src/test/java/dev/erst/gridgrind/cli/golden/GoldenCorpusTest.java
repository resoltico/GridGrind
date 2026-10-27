package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.discovery.Recipe;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestFactory;
import org.junit.jupiter.api.io.TempDir;
import org.opentest4j.AssertionFailedError;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ObjectNode;

/** Behavioural observations made before changing the model, wire vocabulary, or hosts. */
class GoldenCorpusTest {
  @TempDir Path temporary;

  @TestFactory
  Stream<DynamicTest> examplesAndRecipes() throws IOException {
    List<DynamicTest> tests = new ArrayList<>();
    for (Recipe recipe : RecipeRegistry.recipes()) {
      for (GoldenLauncher launcher : GoldenLauncher.values()) {
        tests.add(
            DynamicTest.dynamicTest(
                "recipe/" + recipe.id() + "/" + launcher, () -> recipe(recipe, launcher)));
      }
    }
    try (Stream<Path> files = Files.list(GoldenLauncher.REPOSITORY.resolve("examples"))) {
      for (Path file : files.filter(path -> path.toString().endsWith(".json")).sorted().toList()) {
        Recipe recipe =
            RecipeRegistry.recipes().stream()
                .filter(value -> value.requestFileName().equals(file.getFileName().toString()))
                .findFirst()
                .orElseThrow();
        for (GoldenLauncher launcher : GoldenLauncher.values()) {
          tests.add(
              DynamicTest.dynamicTest(
                  "example/" + file.getFileName() + "/" + launcher,
                  () -> example(file, recipe, launcher)));
        }
      }
    }
    assertFalse(tests.isEmpty(), "The golden execution corpus must not be empty");
    return tests.stream();
  }

  @TestFactory
  Stream<DynamicTest> discovery() {
    List<DiscoveryCase> commands =
        List.of(
            new DiscoveryCase("discovery/0-help", List.of("help")),
            new DiscoveryCase("discovery/1-help-protocol", List.of("help", "protocol")),
            new DiscoveryCase("discovery/2-help-guidance", List.of("help", "guidance")),
            new DiscoveryCase("discovery/3-help", List.of("help", "--format", "structured")),
            new DiscoveryCase("discovery/4-version", List.of("version")),
            new DiscoveryCase("discovery/5-version", List.of("version", "--format", "structured")),
            new DiscoveryCase("discovery/6-license", List.of("license")),
            new DiscoveryCase("discovery/7-license", List.of("license", "--format", "structured")),
            new DiscoveryCase("discovery/8-print-request-template", List.of("template", "request")),
            new DiscoveryCase("discovery/9-print-grant-template", List.of("template", "grant")),
            new DiscoveryCase("discovery/10-print-protocol-catalog", List.of("catalog")),
            new DiscoveryCase(
                "discovery/11-print-protocol-catalog", List.of("catalog", "search", "cell text")),
            new DiscoveryCase(
                "discovery/12-print-protocol-catalog", List.of("catalog", "show", "SET_CELL")),
            new DiscoveryCase("discovery/13-print-recipe-catalog", List.of("recipe", "list")),
            new DiscoveryCase(
                "discovery/14-print-recipe-keyword-match",
                List.of("recipe", "match", "budget chart report")),
            new DiscoveryCase("discovery/15-schema-request", List.of("schema", "request")),
            new DiscoveryCase("discovery/16-schema-grant", List.of("schema", "grant")),
            new DiscoveryCase("discovery/17-schema-response", List.of("schema", "response")));
    List<DynamicTest> tests = new ArrayList<>();
    for (DiscoveryCase command : commands) {
      List<String> arguments = command.arguments();
      String id = command.id();
      for (GoldenLauncher launcher : GoldenLauncher.values()) {
        tests.add(
            DynamicTest.dynamicTest(
                id + "/" + launcher,
                () -> {
                  Path workspace = workspace(id, launcher);
                  GoldenLauncher.Result result =
                      launcher.run(workspace, arguments.toArray(String[]::new));
                  assertEquals(0, result.exit(), () -> result.stdout() + result.stderr());
                  GoldenSnapshots.verify(
                      id, launcher, workspace, GoldenSnapshots.observation(result));
                }));
      }
    }
    return tests.stream();
  }

  @TestFactory
  Stream<DynamicTest> doctorProblems() throws IOException {
    List<DynamicTest> tests = new ArrayList<>();
    Path directory = GoldenLauncher.REPOSITORY.resolve("cli/src/test/resources/golden/problems");
    try (Stream<Path> files = Files.list(directory)) {
      for (Path file : files.filter(Files::isRegularFile).sorted().toList()) {
        for (GoldenLauncher launcher : GoldenLauncher.values()) {
          String id = "doctor/" + file.getFileName();
          tests.add(
              DynamicTest.dynamicTest(
                  id + "/" + launcher,
                  () -> {
                    Path workspace = workspace(id, launcher);
                    Path request = workspace.resolve("request.json");
                    Files.copy(file, request);
                    GoldenLauncher.Result result =
                        launcher.run(workspace, "doctor", "--request", request.toString());
                    assertEquals(1, result.exit(), () -> result.stdout() + result.stderr());
                    JsonNode report = GoldenSnapshots.JSON.readTree(result.stdout());
                    assertFalse(
                        report.path("valid").asBoolean(true), "Doctor must reject this input");
                    GoldenSnapshots.verify(
                        id, launcher, workspace, GoldenSnapshots.observation(result));
                  }));
        }
      }
    }
    assertTrue(tests.size() >= 20, "The doctor corpus must retain all intake defect families");
    return tests.stream();
  }

  @Test
  void oneCharacterSetCellChangeIsDetected() throws IOException, InterruptedException {
    for (GoldenLauncher launcher : GoldenLauncher.values()) {
      Path workspace = workspace("negative-control", launcher);
      Path request = workspace.resolve("request.json");
      String original =
          Files.readString(
              GoldenLauncher.REPOSITORY.resolve("cli/src/test/resources/golden/control.json"));
      Files.writeString(request, original);
      ObjectNode baseline = GoldenWorkbook.execute(launcher, workspace, request);
      GoldenSnapshots.verify("control", launcher, workspace, baseline);
      String changed = original.replace("alpha", "alphb");
      assertFalse(original.equals(changed), "Negative control must change its SET_CELL operand");
      Files.writeString(request, changed);
      ObjectNode observation = GoldenWorkbook.execute(launcher, workspace, request);
      String actual = GoldenSnapshots.canonical(observation, workspace);
      assertTrue(actual.contains("alphb"), "The changed cell must be observed after reopening");
      assertThrows(
          AssertionFailedError.class,
          () ->
              GoldenSnapshots.assertMatches(
                  GoldenLauncher.REPOSITORY.resolve(
                      "cli/src/test/resources/golden/snapshots/control.json"),
                  actual));
    }
  }

  @TestFactory
  Stream<DynamicTest> executionFailures() throws IOException {
    List<DynamicTest> tests = new ArrayList<>();
    Path directory = GoldenLauncher.REPOSITORY.resolve("cli/src/test/resources/golden/errors");
    try (Stream<Path> files = Files.list(directory)) {
      for (Path file : files.filter(Files::isRegularFile).sorted().toList()) {
        for (GoldenLauncher launcher : GoldenLauncher.values()) {
          String id = "errors/" + file.getFileName();
          tests.add(
              DynamicTest.dynamicTest(
                  id + "/" + launcher,
                  () -> {
                    Path workspace = workspace(id, launcher);
                    JsonNode fixture = GoldenSnapshots.JSON.readTree(Files.readAllBytes(file));
                    Path request = workspace.resolve("request.json");
                    Files.writeString(
                        request, GoldenSnapshots.JSON.writeValueAsString(fixture.path("request")));
                    Path grant =
                        GoldenWorkbook.grant(fixture.path("request"), workspace, "grant.json");
                    GoldenLauncher.Result result =
                        launcher.run(
                            workspace,
                            "run",
                            "--request",
                            request.toString(),
                            "--grant",
                            grant.toString());
                    assertEquals(fixture.path("exit").asInt(), result.exit(), result.stdout());
                    assertTrue(
                        result.stdout().contains(fixture.path("code").asString()), result.stdout());
                    GoldenSnapshots.verify(
                        id, launcher, workspace, GoldenSnapshots.observation(result));
                  }));
        }
      }
    }
    assertFalse(tests.isEmpty(), "Execution failure snapshots must be discovered");
    return tests.stream();
  }

  private void recipe(Recipe recipe, GoldenLauncher launcher)
      throws IOException, InterruptedException {
    String id = "recipes/" + recipe.id();
    Path workspace = workspace(id, launcher);
    GoldenLauncher.Result printed = launcher.run(workspace, "recipe", "show", recipe.id());
    assertEquals(recipe.requiredWorkspacePaths().isEmpty() ? 0 : 2, printed.exit());
    ObjectNode observation = GoldenSnapshots.JSON.createObjectNode();
    observation.set("print", GoldenSnapshots.observation(printed));
    Path materialized = materialize(recipe, launcher, workspace);
    GoldenLauncher.Result detail =
        launcher.run(workspace, "recipe", "show", recipe.id(), "--details");
    assertEquals(0, detail.exit(), detail.stdout());
    observation.set("detail", GoldenSnapshots.observation(detail));
    observation.set(
        "materializedRequest",
        GoldenSnapshots.JSON.readTree(
            Files.readAllBytes(materialized.resolve(recipe.requestFileName()))));
    Path request = materialized.resolve(recipe.requestFileName());
    ObjectNode execution = GoldenWorkbook.execute(launcher, materialized, request);
    observation.set("execution", execution);
    GoldenSnapshots.verify(
        id,
        launcher,
        workspace,
        observation,
        GoldenSignaturePreviews.certify(request, materialized, execution));
  }

  private void example(Path file, Recipe recipe, GoldenLauncher launcher)
      throws IOException, InterruptedException {
    String id = "examples/" + file.getFileName();
    Path workspace = workspace(id, launcher);
    Path materialized = materialize(recipe, launcher, workspace);
    Path request = materialized.resolve(recipe.requestFileName());
    Files.write(request, Files.readAllBytes(file));
    ObjectNode execution = GoldenWorkbook.execute(launcher, materialized, request);
    GoldenSnapshots.verify(
        id,
        launcher,
        workspace,
        execution,
        GoldenSignaturePreviews.certify(request, materialized, execution));
  }

  private static Path materialize(Recipe recipe, GoldenLauncher launcher, Path workspace)
      throws IOException, InterruptedException {
    Path materialized = workspace.resolve("recipe");
    GoldenLauncher.Result result =
        launcher.run(
            workspace,
            "recipe",
            "materialize",
            recipe.id(),
            "--workspace",
            materialized.toString());
    assertEquals(0, result.exit(), () -> result.stdout() + result.stderr());
    assertEquals(
        GoldenSnapshots.JSON.readTree(result.stdout()),
        GoldenSnapshots.JSON.readTree(
            Files.readAllBytes(materialized.resolve(recipe.requestFileName()))));
    for (String asset : recipe.requiredWorkspacePaths()) {
      assertTrue(
          Files.isRegularFile(materialized.resolve(asset)), () -> "Missing recipe asset: " + asset);
    }
    if ("PACKAGE_SECURITY_INSPECTION".equals(recipe.id())) {
      Path secrets = Files.createDirectories(materialized.resolve("fixture-secrets"));
      Files.writeString(
          secrets.resolve("source-open-password"), "GridGrind-2026", StandardCharsets.UTF_8);
    }
    return materialized;
  }

  private Path workspace(String id, GoldenLauncher launcher) throws IOException {
    return Files.createDirectories(temporary.resolve(id + "-" + launcher));
  }

  /** Stable observations retain their artifact IDs while command grammar changes. */
  private record DiscoveryCase(String id, List<String> arguments) {}
}
