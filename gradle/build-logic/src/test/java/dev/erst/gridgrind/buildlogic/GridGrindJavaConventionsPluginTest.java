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

/** Functional contract for the shared Java and JaCoCo convention plugin. */
class GridGrindJavaConventionsPluginTest {
  @TempDir Path projectDirectory;

  @Test
  void writesJacocoEvidenceUnderTheOwningProjectsBuildDirectory() throws IOException {
    configureFixture();

    BuildResult result = runner("printJacocoDestination").build();

    Path expected = projectDirectory.toRealPath().resolve("build/jacoco/test.exec");
    assertTrue(result.getOutput().contains("jacocoDestination=" + expected));
    assertFalse(result.getOutput().contains("gridgrind-jacoco"));
  }

  @Test
  void incrementalCompilationRetainsCleanDiagnosticsAndRemovesRetiredClasses() throws IOException {
    configureFixture();
    Path sources = projectDirectory.resolve("src/main/java/fixture");
    Files.createDirectories(sources);
    Files.writeString(
        sources.resolve("package-info.java"),
        "@org.jspecify.annotations.NullMarked package fixture;\n");
    Files.writeString(
        sources.resolve("Service.java"),
        "package fixture; public class Service { public static String value() { return \"ok\"; }"
            + " }\n");
    Files.writeString(
        sources.resolve("Usage.java"),
        "package fixture; public class Usage { public int length() { return"
            + " Service.value().length(); } }\n");
    Path retired = sources.resolve("Old.java");
    Files.writeString(
        retired,
        "package fixture; public sealed interface Old permits Old.Nested { final class Nested"
            + " implements Old {} }\n");
    runner("compileJava", "--build-cache").build();
    assertEquals(
        TaskOutcome.UP_TO_DATE,
        runner("compileJava", "--build-cache").build().task(":compileJava").getOutcome());
    Files.delete(retired);
    Path moved = projectDirectory.resolve("src/main/java/moved/Replacement.java");
    Files.createDirectories(moved.getParent());
    Files.writeString(moved, "package moved; public class Replacement {}\n");
    runner("compileJava", "--build-cache").build();
    assertFalse(
        Files.exists(projectDirectory.resolve("build/classes/java/main/fixture/Old.class")));
    assertFalse(
        Files.exists(projectDirectory.resolve("build/classes/java/main/fixture/Old$Nested.class")));
    assertTrue(
        Files.exists(projectDirectory.resolve("build/classes/java/main/moved/Replacement.class")));
    Files.writeString(
        sources.resolve("Service.java"),
        "package fixture; public class Service { public static @org.jspecify.annotations.Nullable"
            + " String value() { return null; } }\n");
    var incrementalFailure = runner("compileJava").buildAndFail();
    var cleanFailure = runner("clean", "compileJava").buildAndFail();
    assertTrue(
        incrementalFailure.getOutput().contains("[NullAway]"), incrementalFailure.getOutput());
    assertTrue(cleanFailure.getOutput().contains("[NullAway]"), cleanFailure.getOutput());
    Files.writeString(
        sources.resolve("Usage.java"),
        "package fixture; public class Usage { public int length() { String value ="
            + " Service.value(); return value == null ? 0 : value.length(); } }\n");
    runner("compileJava", "--build-cache").build();
    Files.writeString(
        sources.resolve("Usage.java"),
        "package fixture; public class Usage { public boolean same(String a, String b) { return a"
            + " == b; } }\n");
    assertTrue(runner("compileJava").buildAndFail().getOutput().contains("[ReferenceEquality]"));
    Files.writeString(
        sources.resolve("Usage.java"),
        "package fixture; public class Usage { public boolean same(String a, String b) { return"
            + " a.equals(b); } }\n");
    runner("compileJava", "--build-cache").build();
    runner("clean").build();
    assertEquals(
        TaskOutcome.FROM_CACHE,
        runner("compileJava", "--build-cache").build().task(":compileJava").getOutcome());
    assertFalse(
        Files.exists(projectDirectory.resolve("build/classes/java/main/fixture/Old.class")));
  }

  @Test
  void changedClasspathNullabilityMatchesACleanDependentCompile() throws IOException {
    configureFixture();
    Files.writeString(
        projectDirectory.resolve("settings.gradle.kts"),
        "rootProject.name = \"fixture\"\ninclude(\"library\")\n");
    Files.writeString(
        projectDirectory.resolve("build.gradle"),
        "\ndependencies { implementation project(':library') }\n",
        java.nio.file.StandardOpenOption.APPEND);
    Path library = projectDirectory.resolve("library");
    Files.createDirectories(library);
    Files.writeString(
        library.resolve("build.gradle"),
        "plugins { id 'java-library'; id 'gridgrind.java-conventions' }\n");
    Path librarySources = library.resolve("src/main/java/fixture");
    Path usageSources = projectDirectory.resolve("src/main/java/fixture");
    Files.createDirectories(librarySources);
    Files.createDirectories(usageSources);
    for (Path directory : java.util.List.of(librarySources, usageSources)) {
      Files.writeString(
          directory.resolve("package-info.java"),
          "@org.jspecify.annotations.NullMarked package fixture;\n");
    }
    Path api = librarySources.resolve("Library.java");
    Files.writeString(
        api,
        "package fixture; public class Library { public static String value() { return \"ok\"; }"
            + " }\n");
    Files.writeString(
        usageSources.resolve("Usage.java"),
        "package fixture; public class Usage { public int length() { return"
            + " Library.value().length(); } }\n");
    runner(":compileJava").build();
    Files.writeString(
        api,
        "package fixture; public class Library { public static @org.jspecify.annotations.Nullable"
            + " String value() { return null; } }\n");
    var incremental = runner(":compileJava").buildAndFail();
    var clean = runner("clean", ":library:clean", ":compileJava").buildAndFail();
    assertTrue(incremental.getOutput().contains("[NullAway]"), incremental.getOutput());
    assertTrue(clean.getOutput().contains("[NullAway]"), clean.getOutput());
  }

  private void configureFixture() throws IOException {
    Files.createDirectories(projectDirectory.resolve("gradle/pmd"));
    Files.writeString(
        projectDirectory.resolve("settings.gradle.kts"), "rootProject.name = \"fixture\"\n");
    Files.writeString(
        projectDirectory.resolve("gradle.properties"),
        "gridgrindJavaVersion=27\ngridgrindJavaTargetVersion=26\n");
    Files.writeString(projectDirectory.resolve("gradle/pmd/ruleset.xml"), minimalRuleset());
    Files.writeString(projectDirectory.resolve("gradle/pmd/test-ruleset.xml"), minimalRuleset());
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
  }

  private static String minimalRuleset() {
    return """
    <?xml version="1.0"?>
    <ruleset name="fixture" xmlns="http://pmd.sourceforge.net/ruleset/2.0.0" />
    """;
  }

  private static String versionCatalog() {
    try {
      return Files.readString(
          Path.of(System.getProperty("gridgrind.repository.root"))
              .resolve("gradle/libs.versions.toml"));
    } catch (IOException exception) {
      throw new java.io.UncheckedIOException(exception);
    }
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(projectDirectory.toFile())
        .withPluginClasspath()
        .withArguments(arguments)
        .forwardOutput();
  }
}
