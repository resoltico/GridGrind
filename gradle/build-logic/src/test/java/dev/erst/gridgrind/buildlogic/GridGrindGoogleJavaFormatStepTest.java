package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.testkit.runner.GradleRunner;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises current formatter imports, layout, idempotence and configuration-cache reuse. */
class GridGrindGoogleJavaFormatStepTest {
  @TempDir Path root;

  @Test
  void formatsAndChecksWithTheCurrentPublicFormatterApi() throws Exception {
    Files.writeString(root.resolve("settings.gradle"), "rootProject.name='formatter-fixture'\n");
    Path repository = Path.of(System.getProperty("gridgrind.repository.root"));
    String jvmArgs = Files.readString(repository.resolve("gradle.properties")).lines()
        .filter(line -> line.startsWith("org.gradle.jvmargs=")).findFirst().orElseThrow();
    Files.writeString(root.resolve("gradle.properties"), jvmArgs + "\n");
    Files.writeString(root.resolve("build.gradle"), """
        plugins { id 'com.diffplug.spotless' }
        spotless {
          java {
            target 'src/**/*.java'
            addStep(dev.erst.gridgrind.buildlogic.GridGrindGoogleJavaFormatStep.create('1.37.0'))
            formatAnnotations()
          }
        }
        """);
    Path source = root.resolve("src/Sample.java");
    Files.createDirectories(source.getParent());
    Files.writeString(source,
        "import java.util.Set;\nimport java.util.Map;\nimport java.util.List;\nclass Sample{Map<String,List<String>> names=Map.of(\"x\",List.of(\"x\"));}\n");
    runner("spotlessApply", "--configuration-cache").build();
    String expected = "import java.util.List;\nimport java.util.Map;\n\nclass Sample {\n  Map<String, List<String>> names = Map.of(\"x\", List.of(\"x\"));\n}\n";
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
    assertTrue(Files.readString(source).contains("\"" + literal + "\""),
        "Formatting must not split long string literals");
    runner("spotlessCheck", "--configuration-cache").build();
    Files.writeString(source, "class Broken {\n");
    assertTrue(runner("spotlessCheck", "--configuration-cache").buildAndFail()
        .getOutput().contains("spotlessJava"));
  }

  private GradleRunner runner(String... arguments) {
    return GradleRunner.create().withProjectDir(root.toFile()).withPluginClasspath()
        .withArguments(arguments);
  }
}
