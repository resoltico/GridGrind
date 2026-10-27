package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.render.ProductInfo;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.io.TempDir;

/**
 * Actual image, legal, published-recipe and backend checks; explicitly selected by containerTest.
 */
@Tag("container")
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class ContainerTest {
  private Path workspace;
  private DockerFixture docker;

  @BeforeAll
  void buildImage(@TempDir Path directory) throws Exception {
    workspace = directory;
    docker = new DockerFixture(workspace);
    verifyPinnedPlatforms();
    docker.build(
        GoldenLauncher.REPOSITORY,
        List.of("--build-arg", "GRIDGRIND_VERSION=" + ProductInfo.version()));
  }

  @AfterAll
  void removeOwnedResources() throws Exception {
    if (docker != null) docker.close();
  }

  @Test
  void packagedIdentityLegalFilesAndPinnedPlatformsArePresent() throws Exception {
    var help = docker.runWorkbook(workspace, "help");
    assertTrue(DockerFixture.assertSuccess(help).contains("Usage:"));
    assertEquals("", help.stderr());
    var version =
        GoldenSnapshots.JSON.readTree(
            DockerFixture.assertSuccess(
                docker.runWorkbook(workspace, "version", "--format", "structured")));
    assertEquals(ProductInfo.version(), version.path("version").asText());
    assertTrue(
        DockerFixture.assertSuccess(docker.runWorkbook(workspace, "license"))
            .contains("Mozilla Public License"));
    assertEquals(
        "GridGrind",
        DockerFixture.assertSuccess(
                docker.command(
                    60,
                    "image",
                    "inspect",
                    docker.image(),
                    "--format",
                    "{{index .Config.Labels \"org.opencontainers.image.title\"}}"))
            .strip());
    String licenseLabel =
        DockerFixture.assertSuccess(
                docker.command(
                    60,
                    "image",
                    "inspect",
                    docker.image(),
                    "--format",
                    "{{index .Config.Labels \"org.opencontainers.image.licenses\"}}"))
            .strip();
    assertTrue(licenseLabel.isEmpty() || "<no value>".equals(licenseLabel));
    DockerFixture.assertSuccess(
        docker.run(
            List.of("--entrypoint", "sh"),
            "-c",
            "for f in LICENSE NOTICE PATENTS.md LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE; do test -s /usr/share/doc/gridgrind/$f || exit 1; done; test -f /usr/share/doc/base-files/copyright; find /usr/lib/jvm -name LICENSE -type f -print -quit | grep -q ."));
  }

  @SuppressWarnings(
      "PMD.AvoidInstantiatingObjectsInLoops") // Each base image needs an independent platform set.
  private void verifyPinnedPlatforms() throws java.io.IOException, InterruptedException {
    for (String line : Files.readAllLines(GoldenLauncher.REPOSITORY.resolve("Dockerfile"))) {
      if (!line.startsWith("FROM ")) continue;
      var manifest =
          GoldenSnapshots.JSON.readTree(
              DockerFixture.assertSuccess(
                  docker.command(
                      120,
                      "buildx",
                      "imagetools",
                      "inspect",
                      "--raw",
                      line.substring(5).strip().replaceFirst("\\s.*$", ""))));
      Set<String> platforms = new HashSet<>();
      for (var entry : manifest.path("manifests")) {
        platforms.add(
            entry.path("platform").path("os").asText()
                + "/"
                + entry.path("platform").path("architecture").asText());
      }
      assertTrue(platforms.containsAll(Set.of("linux/amd64", "linux/arm64")), platforms.toString());
    }
  }

  @Test
  void everyPublishedRecipeMaterializesExecutesAndReopens() throws Exception {
    for (var recipe : RecipeRegistry.recipes()) {
      Path root = Files.createDirectories(workspace.resolve("recipes").resolve(recipe.id()));
      Path materialized = root.resolve("published workspace");
      DockerFixture.assertSuccess(
          docker.runWorkbook(
              root, "recipe", "materialize", recipe.id(), "--workspace", materialized.toString()));
      if ("PACKAGE_SECURITY_INSPECTION".equals(recipe.id())) {
        Path secrets = Files.createDirectories(materialized.resolve("fixture-secrets"));
        Files.writeString(secrets.resolve("source-open-password"), "GridGrind-2026");
      }
      GoldenWorkbook.execute(
          docker::runWorkbook, materialized, materialized.resolve(recipe.requestFileName()));
    }
  }

  @Test
  void customWorkdirProgressFontsStreamingAndReadbackWork() throws Exception {
    ContainerScenarios.verify(
        docker::runWorkbook, Files.createDirectories(workspace.resolve("smoke paths [#]")));
  }
}
