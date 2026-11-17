package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.GradleRunner;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises current formatter imports, layout, idempotence and configuration-cache reuse. */
class BuiltInGoogleJavaFormatTest {
  @TempDir Path root;

  @Test
  void formatsAndChecksWithTheCurrentPublicFormatterApi() throws Exception {
    Files.writeString(
        root.resolve("settings.gradle.kts"), "rootProject.name = \"formatter-fixture\"\n");
    // A fresh daemon without inherited javac exports exercises Spotless's own integration.
    Files.writeString(
        root.resolve("gradle.properties"),
        "org.gradle.jvmargs=-Xmx256m\ngridgrindJavaVersion=27\ngridgrindJavaTargetVersion=26\n");
    Path repository = Path.of(System.getProperty("gridgrind.repository.root"));
    Files.createDirectories(root.resolve("gradle/pmd"));
    Files.copy(
        repository.resolve("gradle/libs.versions.toml"), root.resolve("gradle/libs.versions.toml"));
    for (String name : java.util.List.of("ruleset.xml", "test-ruleset.xml")) {
      Files.writeString(
          root.resolve("gradle/pmd/" + name),
          "<ruleset name=\"fixture\" xmlns=\"http://pmd.sourceforge.net/ruleset/2.0.0\" />");
    }
    Files.writeString(
        root.resolve("build.gradle"),
        """
        plugins { id 'java'; id 'gridgrind.java-conventions' }
        repositories { mavenCentral() }
        def catalog = extensions.getByType(org.gradle.api.artifacts.VersionCatalogsExtension).named('libs')
        def requiredLibraries = ['guava', 'errorprone-annotations', 'jspecify'].collect { name ->
          def dependency = catalog.findLibrary(name).get().get()
          dependency.module.name + '-' + dependency.versionConstraint.requiredVersion + '.jar'
        }
        configurations.create('formatterProbe')
        dependencies.add('formatterProbe', 'com.google.googlejavaformat:google-java-format:' +
            catalog.findVersion('google-java-format').get().requiredVersion)
        def provisionedLibraries = objects.fileCollection().from(configurations.formatterProbe)
        tasks.register('verifyFormatterLibraryVersions') {
          inputs.files(provisionedLibraries)
          doLast {
            def actual = provisionedLibraries.files.collect { it.name }
            requiredLibraries.each { expected -> assert actual.contains(expected): actual }
          }
        }
        spotless {
          java {
            target 'src/**/*.java'
          }
        }
        """);
    Path source = root.resolve("src/Sample.java");
    Files.createDirectories(source.getParent());
    Files.writeString(
        source,
        "import java.util.Set;\n"
            + "import java.util.Map;\n"
            + "import java.util.List;\n"
            + "class Sample{Map<String,List<String>> names=Map.of(\"x\",List.of(\"x\"));}\n");
    runner("verifyFormatterLibraryVersions", "spotlessApply", "--configuration-cache").build();
    String expected =
        "import java.util.List;\n"
            + "import java.util.Map;\n\n"
            + "class Sample {\n"
            + "  Map<String, List<String>> names = Map.of(\"x\", List.of(\"x\"));\n"
            + "}\n";
    assertEquals(expected, Files.readString(source));
    // The first apply changes the file; establish unchanged inputs before checking reuse.
    runner("spotlessApply", "--configuration-cache").build();
    assertEquals(expected, Files.readString(source));
    var repeat = runner("spotlessApply", "--configuration-cache").build();
    assertTrue(repeat.getOutput().contains("Reusing configuration cache."), repeat.getOutput());
    assertEquals(expected, Files.readString(source));
    runner("spotlessCheck", "--configuration-cache").build();
    String literal = "x".repeat(160);
    Files.writeString(source, "class LongString{String text=\"" + literal + "\";}\n");
    runner("spotlessApply", "--configuration-cache").build();
    assertTrue(
        Files.readString(source).contains("\"" + literal + "\""),
        "Formatting must not split long string literals");
    runner("spotlessCheck", "--configuration-cache").build();
    Files.writeString(source, "class Broken {\n");
    assertTrue(
        runner("spotlessCheck", "--configuration-cache")
            .buildAndFail()
            .getOutput()
            .contains("spotlessJava"));
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create()
        .withProjectDir(root.toFile())
        .withPluginClasspath()
        .withArguments(arguments);
  }
}
