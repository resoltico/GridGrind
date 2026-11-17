package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.io.TempDir;

/** Contributor image and cache ownership checks formerly embedded in a shell/Python validator. */
@Tag("devcontainer")
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class DevcontainerTest {
  private DockerFixture docker;

  @BeforeAll
  void buildContributorImage(@TempDir Path directory) throws Exception {
    docker = new DockerFixture(directory);
    docker.build(GoldenLauncher.REPOSITORY, List.of("--target", "contributor"));
  }

  @AfterAll
  void removeOwnedResources() throws Exception {
    if (docker != null) docker.close();
  }

  @Test
  void workspaceConfigurationAndImageProvideTheDeclaredTools() throws Exception {
    var config =
        GoldenSnapshots.JSON.readTree(
            Files.readString(GoldenLauncher.REPOSITORY.resolve(".devcontainer/devcontainer.json")));
    assertTrue(
        config
            .path("features")
            .has(
                "ghcr.io/devcontainers/features/docker-outside-of-docker:1.10.2@sha256:661b0fd4c57a7e0bd14d99701ab9bc824719678a9319118728f8d9c59441b7e4"));
    var feature = config.path("features").properties().iterator().next().getValue();
    assertEquals("none", feature.path("dockerDashComposeVersion").asText());
    assertFalse(feature.path("installDockerBuildx").asBoolean());
    assertFalse(feature.path("installDockerComposeSwitch").asBoolean());
    assertFalse(feature.path("moby").asBoolean());
    assertEquals("vscode", config.path("remoteUser").asText());
    assertEquals("/workspaces/gridgrind", config.path("workspaceFolder").asText());
    assertEquals(
        "./scripts/devcontainer-prepare-user-home.sh", config.path("postStartCommand").asText());
    String mounts = config.path("mounts").toString();
    assertTrue(mounts.contains("target=/home/vscode/.gradle"));
    assertTrue(mounts.contains("target=/home/vscode/.cache"));
    var editor = config.path("customizations").path("vscode");
    java.util.Set<String> installed = new java.util.HashSet<>();
    for (var extension : editor.path("extensions")) installed.add(extension.asText());
    for (String extension :
        List.of("redhat.java", "vscjava.vscode-gradle", "vscjava.vscode-java-test")) {
      assertTrue(installed.contains(extension));
      assertEquals(
          "workspace",
          editor.path("settings").path("remote.extensionKind").path(extension).get(0).asText());
    }
    boolean defaultRuntime = false;
    for (var runtime : editor.path("settings").path("java.configuration.runtimes")) {
      defaultRuntime |=
          "/opt/java/openjdk".equals(runtime.path("path").asText())
              && runtime.path("default").asBoolean();
    }
    assertTrue(defaultRuntime);
    String tools =
        DockerFixture.assertSuccess(
            docker.run(
                List.of(),
                "bash",
                "-lc",
                "set -euo pipefail; . /etc/os-release; test \"$ID\" = ubuntu; test \"$VERSION_ID\""
                    + " = 26.04; java -XshowSettings:properties -version 2>&1 | grep -F"
                    + " 'java.vendor = Eclipse Adoptium'; java --version; javac --version; git"
                    + " --version; python3 --version; shellcheck --version; fc-match 'DejaVu Sans';"
                    + " docker --version; docker buildx version; if command -v docker-compose ||"
                    + " docker compose version >/dev/null 2>&1; then exit 1; fi"));
    var pins = new java.util.Properties();
    try (var input =
        Files.newInputStream(
            GoldenLauncher.REPOSITORY.resolve("release/container-runtime.properties"))) {
      pins.load(input);
    }
    assertTrue(tools.contains("build " + pins.getProperty("TEMURIN_VERSION")), tools);
    assertTrue(tools.contains("Docker version " + pins.getProperty("DOCKER_CLI_VERSION")), tools);
    assertTrue(tools.contains("buildx v" + pins.getProperty("BUILDX_VERSION")), tools);
  }

  @Test
  void mountedCachesBecomeWritableAfterTheActualStartupRepair() throws Exception {
    String gradle = docker.createVolume();
    String cache = docker.createVolume();
    DockerFixture.assertSuccess(
        docker.run(
            List.of(
                "--user",
                "root",
                "--volume",
                gradle + ":/home/vscode/.gradle",
                "--volume",
                cache + ":/home/vscode/.cache"),
            "bash",
            "-lc",
            "set -euo pipefail; install -d -o root -g root /home/vscode/.gradle/wrapper"
                + " /home/vscode/.cache/probe; touch /home/vscode/.gradle/wrapper/root-owned-marker"
                + " /home/vscode/.cache/probe/root-owned-marker"));
    DockerFixture.assertSuccess(
        docker.run(
            List.of(
                "--volume",
                gradle + ":/home/vscode/.gradle",
                "--volume",
                cache + ":/home/vscode/.cache",
                "--volume",
                GoldenLauncher.REPOSITORY + ":/workspace:ro"),
            "bash",
            "-lc",
            "set -euo pipefail; /workspace/scripts/devcontainer-prepare-user-home.sh; touch"
                + " /home/vscode/.gradle/wrapper/user-writable-marker"
                + " /home/vscode/.cache/probe/user-writable-marker"));
  }
}
