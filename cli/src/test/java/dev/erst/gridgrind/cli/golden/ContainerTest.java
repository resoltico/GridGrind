package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.render.ProductInfo;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
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
  void prepareImage(@TempDir Path directory) throws Exception {
    workspace = directory;
    String prebuiltImage = System.getProperty("gridgrind.container.image", "");
    docker =
        prebuiltImage.isBlank()
            ? new DockerFixture(workspace)
            : new DockerFixture(workspace, prebuiltImage);
    verifyPinnedPlatforms();
    if (prebuiltImage.isBlank()) {
      docker.build(
          GoldenLauncher.REPOSITORY,
          List.of("--build-arg", "GRIDGRIND_VERSION=" + ProductInfo.version()));
    }
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
            "set -eu; for f in LICENSE NOTICE PATENTS.md LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE; do test -s /usr/share/doc/gridgrind/$f || exit 1; done; test -f /usr/share/doc/base-files/copyright; test -s \"$JAVA_HOME/README.md\"; test -s \"$JAVA_HOME/LICENSE\"; test -s \"$JAVA_HOME/ADDITIONAL_LICENSE_INFO\"; test -s \"$JAVA_HOME/ASSEMBLY_EXCEPTION\"; test -d \"$JAVA_HOME/legal\"; find \"$JAVA_HOME/legal\" -type f -print -quit | grep -q ."));
    Properties runtime = new Properties();
    try (var input =
        Files.newInputStream(
            GoldenLauncher.REPOSITORY.resolve("release/container-runtime.properties"))) {
      runtime.load(input);
    }
    assertEquals(
        runtime.getProperty("CORRETTO_SOURCE_COMMIT"),
        DockerFixture.assertSuccess(
                docker.run(List.of("--entrypoint", "sh"), "-c", "cat \"$JAVA_HOME/commitId.txt\""))
            .strip());
    assertEquals(
        runtime.getProperty("CORRETTO_VERSION"),
        DockerFixture.assertSuccess(
                docker.run(List.of("--entrypoint", "sh"), "-c", "cat \"$JAVA_HOME/version.txt\""))
            .strip());
  }

  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  // One parser owns stage mappings; every external base needs its own platform set.
  private void verifyPinnedPlatforms() throws java.io.IOException, InterruptedException {
    Properties runtime = new Properties();
    try (var input =
        Files.newInputStream(
            GoldenLauncher.REPOSITORY.resolve("release/container-runtime.properties"))) {
      runtime.load(input);
    }
    Map<String, String> stages = new HashMap<>();
    String finalBase = "";
    for (String line : Files.readAllLines(GoldenLauncher.REPOSITORY.resolve("Dockerfile"))) {
      String[] words = line.strip().split("\\s+", -1);
      if (words.length < 2 || !"FROM".equalsIgnoreCase(words[0])) continue;
      int imageIndex = 1;
      while (imageIndex < words.length && words[imageIndex].startsWith("--")) imageIndex++;
      assertTrue(imageIndex < words.length, "FROM requires an image");
      String base = words[imageIndex];
      boolean localStage =
          stages.containsKey(base.toLowerCase(Locale.ROOT)) || "scratch".equals(base);
      finalBase = stages.getOrDefault(base.toLowerCase(Locale.ROOT), base);
      if (imageIndex + 2 < words.length && "AS".equalsIgnoreCase(words[imageIndex + 1]))
        stages.put(words[imageIndex + 2].toLowerCase(Locale.ROOT), finalBase);
      if (localStage) continue;
      var manifest =
          GoldenSnapshots.JSON.readTree(
              DockerFixture.assertSuccess(
                  docker.command(120, "buildx", "imagetools", "inspect", "--raw", base)));
      Set<String> platforms = new HashSet<>();
      for (var entry : manifest.path("manifests")) {
        platforms.add(
            entry.path("platform").path("os").asText()
                + "/"
                + entry.path("platform").path("architecture").asText());
      }
      assertTrue(platforms.containsAll(Set.of("linux/amd64", "linux/arm64")), platforms.toString());
    }
    assertEquals(
        runtime.getProperty("DEBIAN_BASE_IMAGE"),
        finalBase,
        "Final runtime base must match the source/packaging input pin");
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
