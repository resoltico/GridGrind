package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindWarningCode;
import dev.erst.gridgrind.contract.dto.RequestWarningLocation;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Direct behavioral coverage for request-owned no-follow path capabilities. */
class RequestPathAccessTest {
  @TempDir Path executionRoot;

  @Test
  void materializesAContainedAbsoluteReadAndReportsItsPortabilityWarning() throws Exception {
    Path source = Files.write(executionRoot.resolve("source.txt"), new byte[] {7, 8, 9});

    try (RequestPathAccess access = requestPathAccess()) {
      Path materialized =
          access.materializeRead(source.toString(), "cell text", "gridgrind-test-", ".txt");

      assertArrayEquals(new byte[] {7, 8, 9}, Files.readAllBytes(materialized));
      RequestPathAccess.MaterializedReadIdentity identity =
          access.materializedReadIdentities().getFirst();
      assertEquals("cell text", identity.role());
      assertEquals(source.toString(), identity.path());
      assertEquals(3, identity.byteSize());
      assertEquals(
          "66a6757151f8ee55db127716c7e3dce0be8074b64e20eda542e5c1e46ca9c41e", identity.sha256());
      assertEquals(
          GridGrindWarningCode.NON_PORTABLE_ABSOLUTE_PATH, access.warnings().getFirst().code());
      RequestWarningLocation.RequestPath location =
          assertInstanceOf(
              RequestWarningLocation.RequestPath.class, access.warnings().getFirst().location());
      assertEquals(source.toString(), location.path());
      assertEquals("cell text", location.pathRole());
    }
  }

  @Test
  void rejectsBothRelativeAndAbsoluteEscapePaths() throws Exception {
    Path outside = Files.createTempFile("gridgrind-outside-", ".txt");

    try (RequestPathAccess access = requestPathAccess()) {
      assertThrows(
          RequestPathEscapeException.class,
          () -> access.materializeRead("../outside.txt", "cell text", "gridgrind-test-", ".txt"));
      assertThrows(
          RequestPathEscapeException.class,
          () -> access.materializeRead(outside.toString(), "cell text", "gridgrind-test-", ".txt"));
    } finally {
      Files.deleteIfExists(outside);
    }
  }

  @Test
  void rejectsAContainedSymlinkRatherThanFollowingIt() throws Exception {
    Path target = Files.write(executionRoot.resolve("target.txt"), new byte[] {1});
    Path link = executionRoot.resolve("link.txt");
    Files.createSymbolicLink(link, target.getFileName());

    try (RequestPathAccess access = requestPathAccess()) {
      assertThrows(
          UnsafePathAccessException.class,
          () -> access.materializeRead("link.txt", "cell text", "gridgrind-test-", ".txt"));
    }
  }

  @Test
  void rejectsContainedReadsThatAreOutsideTheHostGrant() throws Exception {
    Files.write(executionRoot.resolve("source.txt"), new byte[] {1});
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    try (RequestPathAccess access =
        new RequestPathAccess(
            executionRoot, (prefix, suffix) -> Files.createTempFile(prefix, suffix), grant)) {
      assertThrows(
          ExecutionAuthorityDeniedException.class,
          () -> access.materializeRead("source.txt", "source", "gridgrind-test-", ".txt"));
    }
  }

  @Test
  void publishesOnlyThroughThePreparedParentAndRejectsObservedTopologyReplacement()
      throws Exception {
    Path outputDirectory = Files.createDirectory(executionRoot.resolve("output"));
    Path staged = Files.write(executionRoot.resolve("staged.xlsx"), new byte[] {3, 2, 1});

    try (RequestPathAccess access = requestPathAccess()) {
      access.prepareOutput(
          "output/result.xlsx", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW);
      access.publishOutput(
          staged,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          new dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome
              .StagedArtifactVerification());
      assertArrayEquals(
          new byte[] {3, 2, 1}, Files.readAllBytes(outputDirectory.resolve("result.xlsx")));
    }

    try (RequestPathAccess access = requestPathAccess()) {
      access.prepareOutput(
          "output/replaced.xlsx", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW);
      Files.move(outputDirectory, executionRoot.resolve("replaced-output"));
      Files.createDirectory(outputDirectory);

      assertThrows(
          WorkbookPublicationException.class,
          () ->
              access.publishOutput(
                  staged,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome
                      .StagedArtifactVerification()));
    }
  }

  @Test
  void rejectsUnpreparedOutputAccessAndConflictingOrExistingOutputBindings() throws Exception {
    Path outputDirectory = Files.createDirectory(executionRoot.resolve("output"));
    Path existing = Files.write(outputDirectory.resolve("existing.xlsx"), new byte[] {1});

    try (RequestPathAccess access = requestPathAccess()) {
      assertEquals(executionRoot, access.executionRoot());
      assertThrows(IllegalStateException.class, access::outputPath);
      assertThrows(
          IllegalStateException.class,
          () ->
              access.publishOutput(
                  existing,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome
                      .StagedArtifactVerification()));
      assertThrows(
          OutputPathAlreadyExistsException.class,
          () ->
              access.prepareOutput(
                  "output/existing.xlsx",
                  "persistence",
                  WorkbookArtifactWriteDisposition.CREATE_NEW));
      access.prepareOutput(
          "output/replaced.xlsx", "persistence", WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
      assertEquals(executionRoot.resolve("output/replaced.xlsx"), access.outputPath());
      access.prepareOutput(
          "output/replaced.xlsx", "persistence", WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
      assertThrows(
          IllegalStateException.class,
          () ->
              access.prepareOutput(
                  "output/other.xlsx",
                  "persistence",
                  WorkbookArtifactWriteDisposition.REPLACE_EXISTING));
    }
  }

  @Test
  void keepsAnAfterPreflightOutputRaceClosed() throws Exception {
    Path outputDirectory = Files.createDirectory(executionRoot.resolve("output"));
    Path staged = Files.write(executionRoot.resolve("staged.xlsx"), new byte[] {3, 2, 1});

    try (RequestPathAccess access = requestPathAccess()) {
      access.prepareOutput(
          "output/race.xlsx", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW);
      Files.write(outputDirectory.resolve("race.xlsx"), new byte[] {9});

      assertThrows(
          WorkbookPublicationException.class,
          () ->
              access.publishOutput(
                  staged,
                  WorkbookArtifactWriteDisposition.CREATE_NEW,
                  new dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PublicationOutcome
                      .StagedArtifactVerification()));
    }
  }

  @Test
  void rejectsAnExecutionRootSymlinkAndAPathThatNamesTheRoot() throws Exception {
    Path target = Files.createDirectory(executionRoot.resolve("target"));
    Path rootLink = executionRoot.resolve("root-link");
    Files.createSymbolicLink(rootLink, target.getFileName());

    try (RequestPathAccess access = requestPathAccess()) {
      assertThrows(
          UnsafePathAccessException.class,
          () -> access.materializeRead(".", "cell text", "gridgrind-test-", ".txt"));
    }
    assertThrows(
        UnsafePathAccessException.class,
        () ->
            new RequestPathAccess(
                    rootLink,
                    (prefix, suffix) -> Files.createTempFile(prefix, suffix),
                    ExecutionGrantTestSupport.noPublication())
                .materializeRead("file.txt", "cell text", "gridgrind-test-", ".txt"));
  }

  @Test
  void preservesThePrimaryCleanupFailureAndSuppressesSubsequentCleanupFailures() {
    java.io.IOException first = new java.io.IOException("first");
    java.io.IOException second = new java.io.IOException("second");

    java.io.IOException failure =
        assertThrows(
            java.io.IOException.class,
            () ->
                RequestPathAccess.closeAll(
                    List.of(
                        () -> {
                          throw first;
                        },
                        () -> {
                          throw second;
                        })));

    assertEquals(first, failure);
    assertEquals(List.of(second), List.of(failure.getSuppressed()));
  }

  @Test
  void doesNotMaskAReadFailureWhenPrivateMaterializationCleanupAlsoFails() throws Exception {
    Files.write(executionRoot.resolve("source.txt"), new byte[] {1});
    Path invalidMaterialization = Files.createDirectory(executionRoot.resolve("not-a-file"));
    Files.write(invalidMaterialization.resolve("child"), new byte[] {2});

    try (RequestPathAccess access =
        new RequestPathAccess(
            executionRoot,
            (prefix, suffix) -> invalidMaterialization,
            ExecutionGrantTestSupport.noPublication(
                List.of(), List.of(executionRoot.resolve("source.txt"))))) {
      assertThrows(
          java.nio.file.FileSystemException.class,
          () -> access.materializeRead("source.txt", "cell text", "gridgrind-test-", ".txt"));
    }

    assertTrue(Files.exists(invalidMaterialization));
  }

  @Test
  void treatsAnUnavailableMandatorySha256ProviderAsAnInternalRuntimeInvariantFailure()
      throws Exception {
    Path source = Files.write(executionRoot.resolve("source.txt"), new byte[] {1});

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                RequestPathDigest.sha256(
                    source,
                    algorithm -> {
                      throw new java.security.NoSuchAlgorithmException(algorithm);
                    }));

    assertEquals("SHA-256 must be available in the Java runtime", failure.getMessage());
  }

  @Test
  void reusesPrivateMaterializationsAndRejectsDirectoryRolesAndUnknownLookupPaths()
      throws Exception {
    Path source = Files.write(executionRoot.resolve("source.txt"), new byte[] {1, 2});
    Files.createDirectory(executionRoot.resolve("source-directory"));
    Files.createDirectory(executionRoot.resolve("output-directory"));

    try (RequestPathAccess access = requestPathAccess()) {
      Path first = access.materializeRead("source.txt", "cell text", "gridgrind-test-", ".txt");
      Path second = access.materializeRead("source.txt", "cell text", "gridgrind-test-", ".txt");
      assertEquals(first, second);
      assertThrows(IllegalStateException.class, () -> access.materializedReadPath("missing.txt"));
      assertThrows(
          SourcePathIsDirectoryException.class,
          () -> access.materializeRead("source-directory", "source", "gridgrind-test-", ".xlsx"));
      assertThrows(
          OutputPathIsDirectoryException.class,
          () ->
              access.prepareOutput(
                  "output-directory", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW));
    }
    assertTrue(Files.exists(source));
  }

  @Test
  void closesSuccessfulPrivateMaterializationsInReverseOwnershipOrder() throws Exception {
    Files.write(executionRoot.resolve("source.txt"), new byte[] {1, 2});
    try (RequestPathAccess access = requestPathAccess()) {
      Path materialized =
          access.materializeRead("source.txt", "cell text", "gridgrind-test-", ".txt");

      assertTrue(Files.isRegularFile(materialized));
      access.close();

      assertTrue(Files.notExists(materialized));
    }
  }

  private RequestPathAccess requestPathAccess() throws java.io.IOException {
    Path tempRoot = Files.createDirectories(executionRoot.resolve("private-temp"));
    return new RequestPathAccess(
        executionRoot,
        (prefix, suffix) -> Files.createTempFile(tempRoot, prefix, suffix),
        ExecutionGrantTestSupport.noPublication(
            List.of(), List.of(executionRoot.resolve("source.txt"))));
  }
}
