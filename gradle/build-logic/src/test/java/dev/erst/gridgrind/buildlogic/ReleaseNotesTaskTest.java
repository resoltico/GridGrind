package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.file.Files;
import java.nio.file.Path;
import org.gradle.api.GradleException;
import org.gradle.testfixtures.ProjectBuilder;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Checks actual release-note extraction, section boundaries and fail-closed version selection. */
class ReleaseNotesTaskTest {
  @TempDir Path root;

  @Test
  void extractsOnlyTheSelectedVersionSection() throws Exception {
    var task = task("## [Unreleased]\nLater\n\n## [1.2.3] - 2026-10-27\n\n### Changed\n- A real change.\n\n## [1.2.2]\nOlder\n");
    task.writeNotes();
    assertEquals("### Changed\n- A real change.\n", Files.readString(root.resolve("notes.md")));
  }

  @Test
  void missingVersionCannotProduceReleaseNotes() throws Exception {
    var task = task("## [Unreleased]\n- A change.\n");
    assertThrows(GradleException.class, task::writeNotes);
  }

  @Test
  void developmentVersionsAndAmbiguousSectionsAreRejected() throws Exception {
    var task = task("## [1.2.3]\nFirst\n## [1.2.3]\nSecond\n");
    assertThrows(GradleException.class, task::writeNotes);
    task.getReleaseVersion().set("0.0.0-dev");
    assertThrows(GradleException.class, task::writeNotes);
  }

  private ReleaseNotesTask task(String text) throws Exception {
    var project = ProjectBuilder.builder().withProjectDir(root.toFile()).build();
    var task = project.getTasks().create("prepareReleaseNotes", ReleaseNotesTask.class);
    task.getReleaseVersion().set("1.2.3");
    task.getChangelog().set(Files.writeString(root.resolve("CHANGELOG.md"), text).toFile());
    task.getNotes().set(root.resolve("notes.md").toFile());
    return task;
  }
}
