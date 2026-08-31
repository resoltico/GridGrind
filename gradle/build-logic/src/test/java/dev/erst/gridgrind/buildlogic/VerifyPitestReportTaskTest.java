package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.BuildResult;
import org.gradle.testkit.runner.GradleRunner;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies that the PIT report task rejects timeout-derived and otherwise non-killed outcomes. */
class VerifyPitestReportTaskTest {
  @TempDir Path projectDirectory;

  @Test
  void acceptsOnlyKilledMutationOutcomes() throws IOException {
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
          targetClasses.set(['example.Target'] as Set)
          targetTests.set(['example.TargetTest'] as Set)
          mutationThreshold.set(100)
          coverageThreshold.set(100)
          testStrengthThreshold.set(100)
          maxSurviving.set(0)
        }

        tasks.named('verifyPitestReport') {
          setDependsOn([])
          mutationReport.set(layout.buildDirectory.file('reports/pitest/mutations.xml'))
          verificationReport.set(layout.buildDirectory.file('reports/pitest/verification.tsv'))
        }
        """);

    writeMutationReport("KILLED");
    runner("verifyPitestReport").build();
    assertTrue(
        Files.readString(projectDirectory.resolve("build/reports/pitest/verification.tsv"))
            .contains("KILLED\t1"));

    writeMutationReport("TIMED_OUT");
    BuildResult failure = runner("verifyPitestReport").buildAndFail();
    assertTrue(failure.getOutput().contains("TIMED_OUT example.Target#execute:12"));
  }

  private void writeMutationReport(String status) throws IOException {
    Path report = projectDirectory.resolve("build/reports/pitest/mutations.xml");
    Files.createDirectories(report.getParent());
    Files.writeString(
        report,
        """
        <mutations>
          <mutation status="%s">
            <mutatedClass>example.Target</mutatedClass>
            <mutatedMethod>execute</mutatedMethod>
            <lineNumber>12</lineNumber>
          </mutation>
        </mutations>
        """.formatted(status));
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(projectDirectory.toFile())
        .withPluginClasspath()
        .withArguments(arguments)
        .forwardOutput();
  }
}
