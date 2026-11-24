package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.gradle.api.GradleException;
import org.gradle.testfixtures.ProjectBuilder;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies normative runtime identity comes from resolution, without replacing vendor metadata. */
class GenerateRuntimeMavenMetadataTaskTest {
  @TempDir Path root;

  @Test
  void supplementsMissingIdentityFromResolutionAndKeepsExistingPublisherMetadataUntouched()
      throws IOException {
    Path missing = jar("misleading-9.9.jar", null, null);
    String genuinePath = "META-INF/maven/example/present/pom.properties";
    byte[] genuine =
        "# publisher-owned comment\ngroupId=example\nartifactId=present\nversion=2.0\n"
            .getBytes(StandardCharsets.ISO_8859_1);
    Path present = jar("present.jar", genuinePath, genuine);
    var task = task("generate");
    task.getComponents()
        .set(
            List.of(
                new RuntimeMavenComponent("example", "missing", "1.2.3", missing.toFile()),
                new RuntimeMavenComponent("example", "present", "2.0", present.toFile())));
    task.generateMetadata();
    Path output = task.getOutputDirectory().get().getAsFile().toPath();
    assertArrayEquals(
        "artifactId=missing\ngroupId=example\nversion=1.2.3\n"
            .getBytes(StandardCharsets.ISO_8859_1),
        Files.readAllBytes(output.resolve("META-INF/maven/example/missing/pom.properties")));
    assertFalse(
        Files.exists(output.resolve(genuinePath)), "Genuine input metadata must not be duplicated");
    try (var jar = new java.util.zip.ZipFile(present.toFile())) {
      assertArrayEquals(genuine, jar.getInputStream(jar.getEntry(genuinePath)).readAllBytes());
    }
  }

  @Test
  void rejectsPublisherMismatchAndUnsafeCoordinatePaths() throws IOException {
    Path bad =
        jar(
            "bad.jar",
            "META-INF/maven/example/library/pom.properties",
            "groupId=example\nartifactId=library\nversion=wrong\n"
                .getBytes(StandardCharsets.ISO_8859_1));
    var mismatch = task("mismatch");
    mismatch
        .getComponents()
        .set(List.of(new RuntimeMavenComponent("example", "library", "1.0", bad.toFile())));
    assertTrue(
        assertThrows(GradleException.class, mismatch::generateMetadata)
            .getMessage()
            .contains("differs from resolution"));
    var escape = task("escape");
    escape
        .getComponents()
        .set(List.of(new RuntimeMavenComponent("../outside", "library", "1.0", bad.toFile())));
    assertTrue(
        assertThrows(GradleException.class, escape::generateMetadata)
            .getMessage()
            .contains("Maven path segment"));
    assertFalse(Files.exists(root.resolve("outside")));
  }

  @Test
  void outputIsRepeatableAndRemovesOnlyStaleGeneratedIdentities() throws IOException {
    Path artifact = jar("library.jar", null, null);
    var task = task("repeat");
    task.getComponents()
        .set(List.of(new RuntimeMavenComponent("example", "library", "1.0", artifact.toFile())));
    task.generateMetadata();
    Path output = task.getOutputDirectory().get().getAsFile().toPath();
    Path metadata = output.resolve("META-INF/maven/example/library/pom.properties");
    byte[] first = Files.readAllBytes(metadata);
    task.generateMetadata();
    assertArrayEquals(first, Files.readAllBytes(metadata));
    Path unrelated = Files.writeString(root.resolve("keep.txt"), "keep");
    task.getComponents().set(List.of());
    task.generateMetadata();
    assertFalse(Files.exists(metadata));
    assertTrue(Files.isRegularFile(unrelated));
  }

  private GenerateRuntimeMavenMetadataTask task(String name) {
    var project = ProjectBuilder.builder().withProjectDir(root.toFile()).build();
    var task = project.getTasks().register(name, GenerateRuntimeMavenMetadataTask.class).get();
    task.getOutputDirectory().set(root.resolve(name + "-output").toFile());
    return task;
  }

  private Path jar(String file, String metadataPath, byte[] metadata) throws IOException {
    Path path = root.resolve(file);
    try (var jar = new ZipOutputStream(Files.newOutputStream(path))) {
      jar.putNextEntry(new ZipEntry("example/Marker.class"));
      jar.write(new byte[] {1, 2, 3});
      jar.closeEntry();
      if (metadataPath != null) {
        jar.putNextEntry(new ZipEntry(metadataPath));
        jar.write(metadata);
        jar.closeEntry();
      }
    }
    return path;
  }
}
