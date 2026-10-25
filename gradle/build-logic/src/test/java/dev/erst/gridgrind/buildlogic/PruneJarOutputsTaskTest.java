package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.gradle.testfixtures.ProjectBuilder;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Protects declared archives and unrelated files while removing stale packaging artifacts. */
class PruneJarOutputsTaskTest {
  @TempDir Path projectDirectory;

  @Test
  void deletesOnlyStaleTopLevelJars() throws IOException {
    Path libraries = Files.createDirectories(projectDirectory.resolve("libs"));
    Path current = Files.writeString(libraries.resolve("current.jar"), "current");
    Path stale = Files.writeString(libraries.resolve("stale.jar"), "stale");
    Path notes = Files.writeString(libraries.resolve("notes.txt"), "notes");
    Path nested = Files.createDirectories(libraries.resolve("nested.jar"));
    Path nestedArchive = Files.writeString(nested.resolve("keep.jar"), "nested");

    PruneJarOutputsTask task = task(libraries);
    task.getExpectedArchiveFileNames().set(List.of("current.jar"));
    task.pruneJarOutputs();

    assertTrue(Files.isRegularFile(current));
    assertFalse(Files.exists(stale));
    assertTrue(Files.isRegularFile(notes));
    assertTrue(Files.isRegularFile(nestedArchive));
  }

  @Test
  void missingOutputDirectoryIsANoOp() {
    Path missing = projectDirectory.resolve("missing");
    assertDoesNotThrow(() -> task(missing).pruneJarOutputs());
    assertFalse(Files.exists(missing));
  }

  private PruneJarOutputsTask task(Path libraries) {
    var project = ProjectBuilder.builder().withProjectDir(projectDirectory.toFile()).build();
    PruneJarOutputsTask task = project.getTasks().create("prune", PruneJarOutputsTask.class);
    task.getLibsDirectory().set(libraries.toFile());
    return task;
  }
}
