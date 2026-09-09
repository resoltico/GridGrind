package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.BuildResult;
import org.gradle.testkit.runner.GradleRunner;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Functional contract for the shared Java and JaCoCo convention plugin. */
class GridGrindJavaConventionsPluginTest {
  @TempDir Path projectDirectory;

  @Test
  void writesJacocoEvidenceUnderTheOwningProjectsBuildDirectory() throws IOException {
    Files.createDirectories(projectDirectory.resolve("gradle/pmd"));
    Files.writeString(
        projectDirectory.resolve("settings.gradle.kts"), "rootProject.name = \"fixture\"\n");
    Files.writeString(
        projectDirectory.resolve("gradle.properties"), "gridgrindJavaVersion=26\n");
    Files.writeString(
        projectDirectory.resolve("gradle/pmd/ruleset.xml"), minimalRuleset());
    Files.writeString(
        projectDirectory.resolve("gradle/pmd/semantic-shape-ruleset.xml"), minimalRuleset());
    Files.writeString(
        projectDirectory.resolve("gradle/pmd/test-ruleset.xml"), minimalRuleset());
    Files.writeString(projectDirectory.resolve("gradle/semantic-shape-policy.tsv"), "");
    Files.writeString(projectDirectory.resolve("gradle/libs.versions.toml"), versionCatalog());
    Files.writeString(
        projectDirectory.resolve("build.gradle"),
        """
        plugins {
          id 'java'
          id 'gridgrind.java-conventions'
        }

        tasks.register('printJacocoDestination') {
          doLast {
            def jacoco = tasks.named('test').get().extensions.getByType(
                org.gradle.testing.jacoco.plugins.JacocoTaskExtension)
            println "jacocoDestination=${jacoco.destinationFile.toPath()}"
          }
        }
        """);

    BuildResult result = runner("printJacocoDestination").build();

    Path expected = projectDirectory.toRealPath().resolve("build/jacoco/test.exec");
    assertTrue(result.getOutput().contains("jacocoDestination=" + expected));
    assertFalse(result.getOutput().contains("gridgrind-jacoco"));
  }

  private static String minimalRuleset() {
    return """
        <?xml version="1.0"?>
        <ruleset name="fixture" xmlns="http://pmd.sourceforge.net/ruleset/2.0.0" />
        """;
  }

  private static String versionCatalog() {
    return """
        [versions]
        errorprone = "2.50.0"
        google-java-format = "1.36.1"
        jacoco = "0.8.15"
        nullaway = "0.14.1"
        pmd = "7.26.0"

        [libraries]
        errorprone-core = { module = "com.google.errorprone:error_prone_core", version.ref = "errorprone" }
        jspecify = { module = "org.jspecify:jspecify", version = "1.0.1" }
        junit-platform-launcher = { module = "org.junit.platform:junit-platform-launcher", version = "6.1.3" }
        nullaway = { module = "com.uber.nullaway:nullaway", version.ref = "nullaway" }
        """;
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(projectDirectory.toFile())
        .withPluginClasspath()
        .withArguments(arguments)
        .forwardOutput();
  }
}
