package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for request-rooted execution binding helpers. */
class CliExecutionBindingsFactoryTest extends GridGrindCliTestSupport {
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
    Path grantPath = writeNoOpGrant();
    Path tempRoot;
    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            Optional.of(requestPath),
            Optional.empty(),
            Optional.empty(),
            Optional.of(grantPath),
            Optional.empty(),
            dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog.requestTemplate(),
            java.io.InputStream.nullInputStream())) {
      tempRoot = bindings.inputs().tempRoot();

      assertTrue(Files.isDirectory(tempRoot));
    }

    assertTrue(Files.notExists(tempRoot));
    Files.deleteIfExists(grantPath);
  }

  @Test
  void managedRequestInputsCleanupRemovesACreatedExplicitParentWhenItWasEmpty() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-cli-bindings-parent-", ".json");
    Path grantPath = writeNoOpGrant();
    Path explicitParent = Path.of("tmp", "gridgrind-cli-created-parent");
    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            Optional.of(requestPath),
            Optional.empty(),
            Optional.of(explicitParent),
            Optional.of(grantPath),
            Optional.empty(),
            dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog.requestTemplate(),
            java.io.InputStream.nullInputStream())) {
      assertTrue(Files.isDirectory(bindings.inputs().tempRoot()));
    }

    assertTrue(Files.notExists(explicitParent.toAbsolutePath().normalize()));
    Files.deleteIfExists(grantPath);
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

  @Test
  void capturesStandardInputForRequestsThatRequireItWithoutASecretProvider() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-standard-input-");
    byte[] standardInput = "Quarterly planning".getBytes(StandardCharsets.UTF_8);

    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            Optional.empty(),
            Optional.of(workspace),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            standardInputRequest(),
            new ByteArrayInputStream(standardInput))) {
      assertArrayEquals(standardInput, bindings.inputs().standardInputBytes().orElseThrow());
      assertTrue(bindings.inputs().secretResolver().isEmpty());
      GridGrindExecutionGrant.Bounded deniedGrant =
          (GridGrindExecutionGrant.Bounded) bindings.inputs().executionGrant();
      assertTrue(deniedGrant.operationIds().isEmpty());
    } finally {
      CliExecutionBindingsFactory.deleteTreeIfExists(workspace);
    }
  }

  @Test
  void capturesStandardInputAndSecretsWhenTheHostSuppliesBothBindings() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-standard-input-secret-");
    Path provider = Files.createDirectory(workspace.resolve("secrets"));
    Files.writeString(provider.resolve("report-password"), "confidential");
    byte[] standardInput = "Quarterly planning".getBytes(StandardCharsets.UTF_8);

    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            Optional.empty(),
            Optional.of(workspace),
            Optional.empty(),
            Optional.empty(),
            Optional.of(provider),
            standardInputRequest(),
            new ByteArrayInputStream(standardInput))) {
      assertArrayEquals(standardInput, bindings.inputs().standardInputBytes().orElseThrow());
      assertArrayEquals(
          "confidential".toCharArray(),
          bindings
              .inputs()
              .secretResolver()
              .orElseThrow()
              .resolve(new dev.erst.gridgrind.contract.dto.SecretReference("report-password")));
    } finally {
      CliExecutionBindingsFactory.deleteTreeIfExists(workspace);
    }
  }

  @Test
  void rejectsUnreadableGrantBeforeCreatingOwnedScratchSpace() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-grant-failure-");
    Path missingGrant = workspace.resolve("missing-grant.json");
    Path scratchParent = workspace.resolve("scratch");

    try {
      CliGrantReadException exception =
          assertThrows(
              CliGrantReadException.class,
              () ->
                  CliExecutionBindingsFactory.create(
                      Optional.empty(),
                      Optional.of(workspace),
                      Optional.of(scratchParent),
                      Optional.of(missingGrant),
                      Optional.empty(),
                      standardInputRequest(),
                      InputStream.nullInputStream()));

      assertEquals(missingGrant, exception.grantPath());
      assertTrue(Files.notExists(scratchParent));
    } finally {
      CliExecutionBindingsFactory.deleteTreeIfExists(workspace);
    }
  }

  @Test
  void rejectsAnUnreadableSecretProviderBeforeCreatingOwnedScratchSpace() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-secret-failure-");
    Path missingProvider = workspace.resolve("missing-provider");
    Path scratchParent = workspace.resolve("scratch");

    try {
      CliSecretReadException exception =
          assertThrows(
              CliSecretReadException.class,
              () ->
                  CliExecutionBindingsFactory.create(
                      Optional.empty(),
                      Optional.of(workspace),
                      Optional.of(scratchParent),
                      Optional.empty(),
                      Optional.of(missingProvider),
                      standardInputRequest(),
                      InputStream.nullInputStream()));

      assertEquals(missingProvider.toString(), exception.referenceId());
      assertTrue(Files.notExists(scratchParent));
    } finally {
      CliExecutionBindingsFactory.deleteTreeIfExists(workspace);
    }
  }

  private static Path writeNoOpGrant() throws IOException {
    Path grantPath = Files.createTempFile("gridgrind-cli-grant-", ".json");
    CliExecutionGrantDocument grant =
        new CliExecutionGrantDocument(
            List.of(),
            List.of(),
            new CliGrantTargetAuthority.WorkbookWide(),
            new CliGrantPublicationAuthority.None(),
            List.of(),
            new CliGrantAcceptancePolicy.MinimumOnly());
    Files.write(grantPath, dev.erst.gridgrind.cli.discovery.GridGrindCliJson.writeBytes(grant));
    return grantPath;
  }

  private static WorkbookPlan standardInputRequest() {
    return GridGrindJson.analyzeRequest(
            minimalRequestJson(
                    "{ \"type\": \"NEW\" }",
                    "{ \"type\": \"NONE\" }",
                    """
                    [
                      {
                        "stepId": "set-title",
                        "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                        "action": {
                          "type": "SET_CELL",
                          "value": { "type": "TEXT", "source": { "type": "STANDARD_INPUT" } }
                        }
                      }
                    ]
                    """)
                .getBytes(StandardCharsets.UTF_8))
        .requireCompletePlan();
  }

  private static void restoreSystemTempProperty(String previous) {
    if (previous == null) {
      System.clearProperty("java.io.tmpdir");
      return;
    }
    System.setProperty("java.io.tmpdir", previous);
  }
}
