package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;
import org.gradle.api.GradleException;
import org.gradle.testfixtures.ProjectBuilder;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies that runtime dependency and NOTICE drift both fail closed. */
class VerifyRuntimeLegalInventoryTaskTest {
  @TempDir Path projectDirectory;

  @Test
  void acceptsExactInventoryAndRejectsArtifactOrNoticeDrift() throws IOException {
    Path artifact = Files.write(projectDirectory.resolve("library-1.0.jar"), new byte[] {0});
    Path notice = Files.writeString(projectDirectory.resolve("NOTICE"), "Library 1.0\n");

    VerifyRuntimeLegalInventoryTask exact = task("verifyExact", artifact, notice);
    exact.getAuditedRuntimeArtifactNames().set(Set.of("library-1.0.jar"));
    exact.getAuditedNoticeMarkers().set(Set.of("Library 1.0"));
    assertDoesNotThrow(exact::verifyInventory);

    VerifyRuntimeLegalInventoryTask artifactDrift = task("verifyArtifactDrift", artifact, notice);
    artifactDrift.getAuditedRuntimeArtifactNames().set(Set.of("library-2.0.jar"));
    artifactDrift.getAuditedNoticeMarkers().set(Set.of("Library 1.0"));
    GradleException artifactFailure =
        assertThrows(GradleException.class, artifactDrift::verifyInventory);
    assertTrue(artifactFailure.getMessage().contains("Runtime legal inventory drifted"));

    VerifyRuntimeLegalInventoryTask noticeDrift = task("verifyNoticeDrift", artifact, notice);
    noticeDrift.getAuditedRuntimeArtifactNames().set(Set.of("library-1.0.jar"));
    noticeDrift.getAuditedNoticeMarkers().set(Set.of("Missing marker"));
    GradleException noticeFailure =
        assertThrows(GradleException.class, noticeDrift::verifyInventory);
    assertTrue(noticeFailure.getMessage().contains("NOTICE is missing audited runtime markers"));
  }

  @Test
  void reportsOneSidedAndDuplicateArtifactDrift() throws IOException {
    Path notice = Files.writeString(projectDirectory.resolve("NOTICE"), "Library 1.0\n");
    Path firstDirectory = Files.createDirectories(projectDirectory.resolve("first"));
    Path secondDirectory = Files.createDirectories(projectDirectory.resolve("second"));
    Path first = Files.write(firstDirectory.resolve("library.jar"), new byte[] {0});
    Path second = Files.write(secondDirectory.resolve("library.jar"), new byte[] {1});

    VerifyRuntimeLegalInventoryTask unaudited = task("verifyUnaudited", first, notice);
    unaudited.getAuditedRuntimeArtifactNames().set(Set.of());
    unaudited.getAuditedNoticeMarkers().set(Set.of());
    GradleException unauditedFailure =
        assertThrows(GradleException.class, unaudited::verifyInventory);
    assertTrue(unauditedFailure.getMessage().contains("Unaudited: [library.jar]"));

    VerifyRuntimeLegalInventoryTask removed = task("verifyRemoved", first, notice);
    removed.getRuntimeArtifacts().setFrom();
    removed.getAuditedRuntimeArtifactNames().set(Set.of("library.jar"));
    removed.getAuditedNoticeMarkers().set(Set.of());
    GradleException removedFailure = assertThrows(GradleException.class, removed::verifyInventory);
    assertTrue(removedFailure.getMessage().contains("No longer resolved: [library.jar]"));

    VerifyRuntimeLegalInventoryTask duplicate = task("verifyDuplicate", first, notice);
    duplicate.getRuntimeArtifacts().from(second.toFile());
    duplicate.getAuditedRuntimeArtifactNames().set(Set.of("library.jar"));
    duplicate.getAuditedNoticeMarkers().set(Set.of());
    GradleException duplicateFailure =
        assertThrows(GradleException.class, duplicate::verifyInventory);
    assertTrue(duplicateFailure.getMessage().contains("duplicate artifact name: library.jar"));
  }

  private VerifyRuntimeLegalInventoryTask task(String name, Path artifact, Path notice) {
    var project = ProjectBuilder.builder().withProjectDir(projectDirectory.toFile()).build();
    VerifyRuntimeLegalInventoryTask task =
        project.getTasks().create(name, VerifyRuntimeLegalInventoryTask.class);
    task.getRuntimeArtifacts().from(artifact.toFile());
    task.getNoticeFile().set(notice.toFile());
    return task;
  }
}
