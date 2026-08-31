package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.BuildResult;
import org.gradle.testkit.runner.GradleRunner;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies that every PIT scope pattern must resolve to compiled production and test classes. */
class VerifyPitestScopeTaskTest {
  @TempDir Path projectDirectory;

  @Test
  void recordsNestedMatchesAndRejectsStalePatterns() throws IOException {
    Files.writeString(projectDirectory.resolve("settings.gradle"), "rootProject.name = 'fixture'\n");
    Files.createDirectories(projectDirectory.resolve("gradle"));
    Files.writeString(
        projectDirectory.resolve("gradle.properties"), "gridgrindJavaVersion=26\n");
    Files.writeString(
        projectDirectory.resolve("gradle/libs.versions.toml"),
        """
        [versions]
        pitest = "1.30.0"
        pitest-junit5-plugin = "1.2.3"
        """);
    Files.writeString(
        projectDirectory.resolve("build.gradle"),
        """
        plugins {
          id 'java'
          id 'gridgrind.mutation-conventions'
        }

        gridgrindMutation {
          targetClasses.set(['example.Target*'] as Set)
          targetTests.set(['example.TargetTest'] as Set)
          mutationThreshold.set(100)
          coverageThreshold.set(100)
          testStrengthThreshold.set(100)
          maxSurviving.set(0)
        }

        tasks.named('verifyPitestScope') {
          setDependsOn([])
          targetClassPatterns.set(['example.Target*'] as Set)
          targetTestPatterns.set(['example.TargetTest'] as Set)
          productionClassDirectories.setFrom(layout.buildDirectory.dir('classes/java/main'))
          testClassDirectories.setFrom(layout.buildDirectory.dir('classes/java/test'))
          reportFile.set(layout.buildDirectory.file('reports/pitest/scope.tsv'))
        }
        """);
    writeClass("build/classes/java/main/example/Target.class");
    writeClass("build/classes/java/main/example/Target$1.class");
    writeClass("build/classes/java/test/example/TargetTest.class");

    runner("verifyPitestScope").build();
    assertTrue(
        Files.readString(projectDirectory.resolve("build/reports/pitest/scope.tsv"))
            .contains("example.Target,example.Target$1"));

    Files.writeString(
        projectDirectory.resolve("build.gradle"),
        Files.readString(projectDirectory.resolve("build.gradle"))
            .replace("example.TargetTest", "example.MissingTest"));
    BuildResult failure = runner("verifyPitestScope").buildAndFail();
    assertTrue(failure.getOutput().contains("example.MissingTest' matched no compiled class"));
  }

  private void writeClass(String relativePath) throws IOException {
    Path classFile = projectDirectory.resolve(relativePath);
    Files.createDirectories(classFile.getParent());
    Files.write(classFile, new byte[] {0});
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(projectDirectory.toFile())
        .withPluginClasspath()
        .withArguments(arguments)
        .forwardOutput();
  }
}
