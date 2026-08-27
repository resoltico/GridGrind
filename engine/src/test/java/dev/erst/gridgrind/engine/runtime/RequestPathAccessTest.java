package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindWarningCode;
import dev.erst.gridgrind.contract.dto.RequestWarningLocation;
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
  void commitsOnlyThroughThePreparedParentAndRejectsObservedTopologyReplacement() throws Exception {
    Path outputDirectory = Files.createDirectory(executionRoot.resolve("output"));
    Path staged = Files.write(executionRoot.resolve("staged.xlsx"), new byte[] {3, 2, 1});

    try (RequestPathAccess access = requestPathAccess()) {
      access.prepareOutput(
          "output/result.xlsx", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW);
      access.commitOutput(staged, WorkbookArtifactWriteDisposition.CREATE_NEW);
      assertArrayEquals(
          new byte[] {3, 2, 1}, Files.readAllBytes(outputDirectory.resolve("result.xlsx")));
    }

    try (RequestPathAccess access = requestPathAccess()) {
      access.prepareOutput(
          "output/replaced.xlsx", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW);
      Files.move(outputDirectory, executionRoot.resolve("replaced-output"));
      Files.createDirectory(outputDirectory);

      assertThrows(
          UnsafePathAccessException.class,
          () -> access.commitOutput(staged, WorkbookArtifactWriteDisposition.CREATE_NEW));
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
          () -> access.commitOutput(existing, WorkbookArtifactWriteDisposition.CREATE_NEW));
      assertThrows(
          OutputPathAlreadyExistsException.class,
          () ->
              access.prepareOutput(
                  "output/existing.xlsx",
                  "persistence",
                  WorkbookArtifactWriteDisposition.CREATE_NEW));
      access.prepareOutput(
          "output/replaced.xlsx", "persistence", WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
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
  void keepsAnAfterPreflightOutputRaceClosedAndClassifiesDirectOutputConflicts() throws Exception {
    Path outputDirectory = Files.createDirectory(executionRoot.resolve("output"));
    Path staged = Files.write(executionRoot.resolve("staged.xlsx"), new byte[] {3, 2, 1});

    try (RequestPathAccess access = requestPathAccess()) {
      access.prepareOutput(
          "output/race.xlsx", "persistence", WorkbookArtifactWriteDisposition.CREATE_NEW);
      Files.write(outputDirectory.resolve("race.xlsx"), new byte[] {9});

      assertThrows(
          UnsafePathAccessException.class,
          () -> access.commitOutput(staged, WorkbookArtifactWriteDisposition.CREATE_NEW));
    }
    assertThrows(
        OutputPathAlreadyExistsException.class,
        () ->
            RequestPathAccess.commitPreparedOutput(
                "output/race.xlsx",
                () -> {
                  throw new java.nio.file.FileAlreadyExistsException("race.xlsx");
                }));
    java.io.IOException ioFailure = new java.io.IOException("disk full");
    assertEquals(
        ioFailure,
        assertThrows(
            java.io.IOException.class,
            () ->
                RequestPathAccess.commitPreparedOutput(
                    "output/race.xlsx",
                    () -> {
                      throw ioFailure;
                    })));
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
                    rootLink, (prefix, suffix) -> Files.createTempFile(prefix, suffix))
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
        new RequestPathAccess(executionRoot, (prefix, suffix) -> invalidMaterialization)) {
      assertThrows(
          java.nio.file.FileSystemException.class,
          () -> access.materializeRead("source.txt", "cell text", "gridgrind-test-", ".txt"));
    }

    assertTrue(Files.exists(invalidMaterialization));
  }

  private RequestPathAccess requestPathAccess() throws java.io.IOException {
    Path tempRoot = Files.createDirectories(executionRoot.resolve("private-temp"));
    return new RequestPathAccess(
        executionRoot, (prefix, suffix) -> Files.createTempFile(tempRoot, prefix, suffix));
  }
}
