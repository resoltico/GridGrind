package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.BuildResult;
import org.gradle.testkit.runner.GradleRunner;
import org.gradle.testkit.runner.TaskOutcome;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Functional contract for the shared mutation convention plugin. */
class GridGrindMutationConventionsPluginTest {
  @TempDir Path projectDirectory;

  @Test
  void wiresAdvisoryPitReportsWithoutDeletingReusableOutput() throws IOException {
    Files.createDirectories(projectDirectory.resolve("gradle"));
    Files.writeString(
        projectDirectory.resolve("settings.gradle"), "rootProject.name = 'fixture'\n");
    Files.writeString(projectDirectory.resolve("gradle.properties"), "gridgrindJavaVersion=26\n");
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
          targetClasses.convention(['example.Target'] as Set)
          targetTests.convention(['example.TargetTest'] as Set)
        }

        tasks.register('printMutationPolicy') {
          doLast {
            def policy = project.extensions.getByName('pitest')
            println "targetClasses=${policy.targetClasses.get()}"
            println "targetTests=${policy.targetTests.get()}"
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

    BuildResult policy = runner("printMutationPolicy", "--max-workers=1").build();
    assertTrue(policy.getOutput().contains("pitestVersion=1.30.0"));
    assertTrue(policy.getOutput().contains("junit5PluginVersion=1.2.3"));
    assertTrue(policy.getOutput().contains("mutators=[STRONGER]"));
    assertTrue(policy.getOutput().contains("mutationThresholdConfigured=false"));
    assertTrue(policy.getOutput().contains("coverageThresholdConfigured=false"));
    assertTrue(policy.getOutput().contains("testStrengthThresholdConfigured=false"));
    assertTrue(policy.getOutput().contains("maxSurvivingConfigured=false"));
    assertTrue(policy.getOutput().contains("failWhenNoMutations=true"));
    assertTrue(
        policy
            .getOutput()
            .contains("jvmArgs=[-Xmx512m, --enable-native-access=ALL-UNNAMED, -Djava.io.tmpdir="));
    assertTrue(policy.getOutput().contains("threads=1"));
    assertTrue(policy.getOutput().contains("timeoutConstInMillis=10000"));
    assertTrue(policy.getOutput().contains("timeoutFactor=3.0"));

    BuildResult dryRun = runner("pitest", "--dry-run").build();
    assertFalse(dryRun.getOutput().contains(":cleanPitestReport SKIPPED"));
    assertTrue(dryRun.getOutput().contains(":pitest SKIPPED"));

    Path staleReport = projectDirectory.resolve("build/reports/pitest/stale.html");
    Files.createDirectories(staleReport.getParent());
    Files.writeString(staleReport, "stale");
    runner("cleanPitestReport").build();
    assertFalse(Files.exists(staleReport));

    BuildResult selected =
        runner(
                "printMutationPolicy",
                "-PmutationClasses=sample.Other",
                "-PmutationTests=sample.OtherTest",
                "-PmutationMutators=MATH")
            .build();
    assertTrue(selected.getOutput().contains("targetClasses=[sample.Other]"));
    assertTrue(selected.getOutput().contains("targetTests=[sample.OtherTest]"));
    assertTrue(selected.getOutput().contains("mutators=[MATH]"));
    assertTrue(
        runner("printMutationPolicy", "-PmutationClasses=,")
            .buildAndFail()
            .getOutput()
            .contains("mutationClasses must contain nonempty"));

    Files.writeString(
        projectDirectory.resolve("build.gradle"),
        """
        plugins {
          id 'java'
          id 'gridgrind.mutation-conventions'
        }
        repositories { mavenCentral() }
        dependencies {
          testImplementation 'org.junit.jupiter:junit-jupiter:6.1.3'
          testRuntimeOnly 'org.junit.platform:junit-platform-launcher:6.1.3'
        }
        pitest {
          targetClasses.convention(['example.Target'] as Set)
          targetTests.convention(['example.TargetTest'] as Set)
          mutators.set(['MATH'] as Set)
        }
        """);
    Path production = projectDirectory.resolve("src/main/java/example/Target.java");
    Path test = projectDirectory.resolve("src/test/java/example/TargetTest.java");
    Files.createDirectories(production.getParent());
    Files.createDirectories(test.getParent());
    Files.writeString(
        production,
        "package example; public class Target { public int increment(int n) { return n + 1; } }");
    Files.writeString(
        test,
        "package example; class TargetTest { @org.junit.jupiter.api.Test void increments() {"
            + " org.junit.jupiter.api.Assertions.assertEquals(2, new Target().increment(1)); } }");
    Files.createDirectories(staleReport.getParent());
    Files.writeString(staleReport, "stale");
    assertEquals(
        TaskOutcome.SUCCESS,
        runner("pitest", "--max-workers=1", "--configuration-cache")
            .build()
            .task(":pitest")
            .getOutcome());
    assertFalse(Files.exists(staleReport));
    assertFalse(Files.exists(projectDirectory.resolve("build/tmp/pitest")));
    assertEquals(
        TaskOutcome.UP_TO_DATE,
        runner("pitest", "--max-workers=1", "--configuration-cache")
            .build()
            .task(":pitest")
            .getOutcome());
    Files.writeString(
        production,
        "package example; public class Target { public int increment(int n) { return n + 2; } }");
    Files.writeString(
        test,
        "package example; class TargetTest { @org.junit.jupiter.api.Test void increments() {"
            + " org.junit.jupiter.api.Assertions.assertEquals(3, new Target().increment(1)); } }");
    assertEquals(
        TaskOutcome.SUCCESS,
        runner("pitest", "--max-workers=1", "--configuration-cache")
            .build()
            .task(":pitest")
            .getOutcome());
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(projectDirectory.toFile())
        .withPluginClasspath()
        .withArguments(arguments)
        .forwardOutput();
  }
}
