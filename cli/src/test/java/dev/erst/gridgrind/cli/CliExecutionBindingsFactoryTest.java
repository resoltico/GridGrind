package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
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
  void tempRootParentDefaultsUnderTheSystemTempDirectory() throws IOException {
    Path systemTemp = Path.of(System.getProperty("java.io.tmpdir")).toAbsolutePath().normalize();

    assertEquals(systemTemp, CliExecutionBindingsFactory.tempRootParent(Optional.empty()));
  }

  @Test
  void createManagedTempRootUsesTheExplicitOverrideWhenPresent() throws IOException {
    Path scratch = Path.of("tmp", "scratch-root");
    CliExecutionBindingsFactory.ManagedTempRoot managedTempRoot =
        CliExecutionBindingsFactory.createManagedTempRoot(Optional.of(scratch));

    try {
      assertEquals(
          scratch.toAbsolutePath().normalize(),
          managedTempRoot.root().getParent().toAbsolutePath().normalize());
      assertTrue(Files.isDirectory(managedTempRoot.root()));
    } finally {
      CliExecutionBindingsFactory.deleteTreeIfExists(managedTempRoot.root());
      managedTempRoot.createdParent().ifPresent(CliExecutionBindingsFactory::deleteTreeIfExists);
    }
  }

  @Test
  void managedRequestInputsCleanupRemovesTheOwnedScratchTree() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-cli-bindings-", ".json");
    Path tempRoot;
    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            Optional.of(requestPath),
            Optional.empty(),
            Optional.empty(),
            dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog.requestTemplate(),
            java.io.InputStream.nullInputStream())) {
      tempRoot = bindings.inputs().tempRoot();

      assertTrue(Files.isDirectory(tempRoot));
    }

    assertTrue(Files.notExists(tempRoot));
  }

  @Test
  void managedRequestInputsCleanupRemovesACreatedExplicitParentWhenItWasEmpty() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-cli-bindings-parent-", ".json");
    Path explicitParent = Path.of("tmp", "gridgrind-cli-created-parent");
    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            Optional.of(requestPath),
            Optional.empty(),
            Optional.of(explicitParent),
            dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog.requestTemplate(),
            java.io.InputStream.nullInputStream())) {
      assertTrue(Files.isDirectory(bindings.inputs().tempRoot()));
    }

    assertTrue(Files.notExists(explicitParent.toAbsolutePath().normalize()));
  }

  @Test
  void deleteTreeIfExistsAllowsNullRoots() {
    assertDoesNotThrow(() -> CliExecutionBindingsFactory.deleteTreeIfExists(null));
  }

  @Test
  void deleteTreeIfExistsSwallowsMissingRoots() {
    Path missing = Path.of("tmp", "gridgrind-missing-root-for-cleanup");
    assertDoesNotThrow(() -> CliExecutionBindingsFactory.deleteTreeIfExists(missing));
  }

  @Test
  void deletePathSwallowsIoFailuresForNonEmptyDirectories() throws Exception {
    Path nonEmptyDirectory = Files.createTempDirectory("gridgrind-cli-non-empty-dir-");
    Files.writeString(nonEmptyDirectory.resolve("child.txt"), "content");

    assertDoesNotThrow(() -> CliExecutionBindingsFactory.deletePath(nonEmptyDirectory));
    assertTrue(Files.exists(nonEmptyDirectory));

    CliExecutionBindingsFactory.deleteTreeIfExists(nonEmptyDirectory);
    assertFalse(Files.exists(nonEmptyDirectory));
  }

  @Test
  void tempRootParentRejectsMissingSystemTempProperty() {
    String previous = System.getProperty("java.io.tmpdir");
    try {
      System.clearProperty("java.io.tmpdir");
      IOException exception =
          assertThrows(
              IOException.class,
              () -> CliExecutionBindingsFactory.tempRootParent(Optional.empty()));

      assertEquals("System temporary-file root is unavailable", exception.getMessage());
    } finally {
      restoreSystemTempProperty(previous);
    }
  }

  @Test
  void tempRootParentRejectsBlankSystemTempProperty() {
    String previous = System.getProperty("java.io.tmpdir");
    try {
      System.setProperty("java.io.tmpdir", "   ");
      IOException exception =
          assertThrows(
              IOException.class,
              () -> CliExecutionBindingsFactory.tempRootParent(Optional.empty()));

      assertEquals("System temporary-file root is unavailable", exception.getMessage());
    } finally {
      restoreSystemTempProperty(previous);
    }
  }

  private static void restoreSystemTempProperty(String previous) {
    if (previous == null) {
      System.clearProperty("java.io.tmpdir");
      return;
    }
    System.setProperty("java.io.tmpdir", previous);
  }
}
