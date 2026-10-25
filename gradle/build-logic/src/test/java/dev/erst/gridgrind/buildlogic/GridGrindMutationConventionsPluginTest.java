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

/** Functional contract for the shared mutation convention plugin. */
class GridGrindMutationConventionsPluginTest {
  @TempDir Path projectDirectory;

  @Test
  void wiresAdvisoryPitReportsAndCleansPreviousOutput() throws IOException {
    Files.createDirectories(projectDirectory.resolve("gradle"));
    Files.writeString(projectDirectory.resolve("settings.gradle"), "rootProject.name = 'fixture'\n");
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

        pitest {
          targetClasses.set(['example.Target'] as Set)
          targetTests.set(['example.TargetTest'] as Set)
        }

        tasks.register('printMutationPolicy') {
          doLast {
            def policy = project.extensions.getByName('pitest')
            println "pitestVersion=${policy.pitestVersion.get()}"
            println "junit5PluginVersion=${policy.junit5PluginVersion.get()}"
            println "mutators=${policy.mutators.get()}"
            println "mutationThresholdConfigured=${policy.mutationThreshold.isPresent()}"
            println "coverageThresholdConfigured=${policy.coverageThreshold.isPresent()}"
            println "testStrengthThresholdConfigured=${policy.testStrengthThreshold.isPresent()}"
            println "maxSurvivingConfigured=${policy.maxSurviving.isPresent()}"
            println "failWhenNoMutations=${policy.failWhenNoMutations.get()}"
            println "jvmArgs=${policy.jvmArgs.get()}"
            println "threads=${policy.threads.get()}"
            println "timeoutConstInMillis=${policy.timeoutConstInMillis.get()}"
            println "timeoutFactor=${policy.timeoutFactor.get()}"
          }
        }
        """);

    BuildResult policy = runner("printMutationPolicy").build();
    assertTrue(policy.getOutput().contains("pitestVersion=1.30.0"));
    assertTrue(policy.getOutput().contains("junit5PluginVersion=1.2.3"));
    assertTrue(policy.getOutput().contains("mutators=[STRONGER]"));
    assertTrue(policy.getOutput().contains("mutationThresholdConfigured=false"));
    assertTrue(policy.getOutput().contains("coverageThresholdConfigured=false"));
    assertTrue(policy.getOutput().contains("testStrengthThresholdConfigured=false"));
    assertTrue(policy.getOutput().contains("maxSurvivingConfigured=false"));
    assertTrue(policy.getOutput().contains("failWhenNoMutations=true"));
    assertTrue(policy.getOutput().contains("jvmArgs=[--enable-native-access=ALL-UNNAMED]"));
    assertTrue(policy.getOutput().contains("threads=4"));
    assertTrue(policy.getOutput().contains("timeoutConstInMillis=10000"));
    assertTrue(policy.getOutput().contains("timeoutFactor=3.0"));

    BuildResult dryRun = runner("pitest", "--dry-run").build();
    assertTrue(dryRun.getOutput().contains(":cleanPitestReport SKIPPED"));
    assertTrue(dryRun.getOutput().contains(":pitest SKIPPED"));

    Path staleReport = projectDirectory.resolve("build/reports/pitest/stale.html");
    Files.createDirectories(staleReport.getParent());
    Files.writeString(staleReport, "stale");
    runner("cleanPitestReport").build();
    assertFalse(Files.exists(staleReport));
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(projectDirectory.toFile())
        .withPluginClasspath()
        .withArguments(arguments)
        .forwardOutput();
  }
}
