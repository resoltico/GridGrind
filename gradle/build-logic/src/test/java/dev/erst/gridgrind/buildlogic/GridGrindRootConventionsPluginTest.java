package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.GradleRunner;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises the owned-build directory gate through its actual cacheable task path. */
class GridGrindRootConventionsPluginTest {
  @TempDir Path root;

  @Test
  void buildSrcGateReusesConfigurationAndRejectsOnlyOwnedBuildDirectories() throws IOException {
    Files.createDirectories(root.resolve("gradle/pmd"));
    Files.createDirectories(root.resolve("gradle/build-logic"));
    Files.createDirectories(root.resolve("module"));
    Files.writeString(
        root.resolve("settings.gradle.kts"),
        """
        rootProject.name = "root-fixture"
        include("module")
        includeBuild("gradle/build-logic") { name = "build-logic" }
        """);
    Files.writeString(
        root.resolve("gradle/build-logic/settings.gradle"), "rootProject.name = 'build-logic'\n");
    Files.writeString(root.resolve("gradle/build-logic/build.gradle"), "tasks.register('test')\n");
    Files.writeString(root.resolve("module/build.gradle"), "plugins { id 'base' }\n");
    Files.writeString(
        root.resolve("build.gradle"), "plugins { id 'gridgrind.root-conventions' }\n");
    Files.writeString(
        root.resolve("gradle.properties"),
        "group=example\nversion=0.0.0-dev\ngridgrindDescription=Fixture\n");
    Path repository = Path.of(System.getProperty("gridgrind.repository.root"));
    Files.copy(
        repository.resolve("gradle/libs.versions.toml"), root.resolve("gradle/libs.versions.toml"));
    for (String name : java.util.List.of("ruleset.xml", "test-ruleset.xml")) {
      Files.writeString(
          root.resolve("gradle/pmd/" + name),
          "<ruleset name=\"fixture\" xmlns=\"http://pmd.sourceforge.net/ruleset/2.0.0\" />");
    }
    // Generated output is not an owned Gradle project and must not broaden this gate's scope.
    Files.createDirectories(root.resolve("build/generated/buildSrc"));
    runner().build();
    var cached = runner().build();
    assertTrue(cached.getOutput().contains("Reusing configuration cache."), cached.getOutput());
    Files.createDirectories(root.resolve("module/buildSrc"));
    var rejected = runner().buildAndFail();
    assertTrue(
        rejected.getOutput().contains("Legacy buildSrc directories are forbidden"),
        rejected.getOutput());
    assertTrue(rejected.getOutput().contains("module/buildSrc"), rejected.getOutput());
    assertTrue(
        rejected.getOutput().contains(":verifyNoLegacyBuildSrc FAILED"), rejected.getOutput());
  }

  private GradleRunner runner() {
    return GradleRunner.create()
        .withProjectDir(root.toFile())
        .withPluginClasspath()
        .withArguments("verifyNoLegacyBuildSrc", "--configuration-cache", "--max-workers=1");
  }
}
