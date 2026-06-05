package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for request-rooted execution binding helpers. */
class CliExecutionBindingsFactoryTest {
  @Test
  void executionWorkingDirectoryPreservesRootPathsThatHaveNoParent() {
    Path root = Path.of("").toAbsolutePath().normalize().getRoot();
    assertNotNull(root);

    assertEquals(
        root,
        CliExecutionBindingsFactory.executionWorkingDirectory(Optional.of(root), Optional.empty()));
  }

  @Test
  void executionWorkingDirectoryUsesExplicitExecutionRootForStdinRequests() {
    Path workspace = Path.of("tmp", "workspace-root");

    assertEquals(
        workspace.toAbsolutePath().normalize(),
        CliExecutionBindingsFactory.executionWorkingDirectory(
            Optional.empty(), Optional.of(workspace)));
  }

  @Test
  void executionWorkingDirectoryRejectsStdinRequestsWithoutAnExplicitExecutionRoot() {
    IllegalArgumentException exception =
        org.junit.jupiter.api.Assertions.assertThrows(
            IllegalArgumentException.class,
            () ->
                CliExecutionBindingsFactory.executionWorkingDirectory(
                    Optional.empty(), Optional.empty()));

    assertEquals("executionRootPath must be present", exception.getMessage());
  }

  @Test
  void tempRootDefaultsUnderTheResolvedWorkingDirectory() {
    Path workspace = Path.of("tmp", "workspace-root").toAbsolutePath().normalize();

    assertEquals(
        workspace.resolve(".gridgrind").resolve("tmp"),
        CliExecutionBindingsFactory.tempRoot(Optional.empty(), workspace));
  }

  @Test
  void tempRootUsesTheExplicitOverrideWhenPresent() {
    Path workspace = Path.of("tmp", "workspace-root").toAbsolutePath().normalize();
    Path scratch = Path.of("tmp", "scratch-root");

    assertEquals(
        scratch.toAbsolutePath().normalize(),
        CliExecutionBindingsFactory.tempRoot(Optional.of(scratch), workspace));
  }
}
