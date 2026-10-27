package dev.erst.gridgrind.fuzz.support;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.io.IOException;
import java.nio.file.ClosedFileSystemException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Checks cleanup of native scratch and failures occurring after workflow storage allocation. */
class FuzzFilesystemTest {
  @TempDir Path directory;

  @Test
  void deletesNestedOwnedFilesWithoutFollowingOutsideLinks() throws IOException {
    Path outside = Files.writeString(directory.resolve("outside"), "keep");
    Path owned = Files.createDirectory(directory.resolve("owned"));
    Files.createSymbolicLink(owned.resolve("link"), outside);
    Files.writeString(Files.createDirectories(owned.resolve("nested")).resolve("file"), "remove");
    FuzzFilesystem.deleteRecursively(owned);
    assertFalse(Files.exists(owned));
    assertEquals("keep", Files.readString(outside));
  }

  @Test
  void workflowConstructionFailureCleansAllocatedStorage() throws IOException {
    Path temporary = Path.of(System.getProperty("java.io.tmpdir"));
    Set<Path> before = workflowDirectories(temporary);
    var data =
        new GridGrindFuzzData() {
          private int bytes;

          public boolean consumeBoolean() {
            return false;
          }

          public byte consumeByte() {
            if (++bytes == 5)
              throw new IllegalArgumentException("injected after storage allocation");
            return 0;
          }

          public int consumeInt(int min, int max) {
            return min;
          }

          public double consumeRegularDouble(double min, double max) {
            return min;
          }

          public int remainingBytes() {
            return 0;
          }
        };
    var failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> OperationSequenceModel.nextProtocolWorkflow(data));
    assertEquals("injected after storage allocation", failure.getMessage());
    assertEquals(before, workflowDirectories(temporary));
  }

  @Test
  void cleanupFailureIsSuppressedOntoTheOriginalFinding() throws IOException {
    Path unavailable = closedFilesystemPath();
    var finding = new IllegalStateException("original invariant failure");
    var actual =
        assertThrows(
            IllegalStateException.class,
            () -> {
              try {
                throw finding;
              } catch (IllegalStateException failure) {
                FuzzFilesystem.cleanupAfterFailure(unavailable, failure);
                throw failure;
              }
            });
    assertSame(finding, actual);
    assertEquals(1, actual.getSuppressed().length);
    assertInstanceOf(ClosedFileSystemException.class, actual.getSuppressed()[0]);
  }

  @Test
  void workflowCleanupAttemptsLaterRootsAfterAnEarlierFailure() throws IOException {
    Path unavailable = closedFilesystemPath();
    Path owned = Files.createDirectories(directory.resolve("later-root"));
    Files.writeString(owned.resolve("payload"), "remove");
    var request =
        ProtocolStepSupport.request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of());
    var workflow = new GeneratedProtocolWorkflow(request, List.of(unavailable, owned));
    assertThrows(ClosedFileSystemException.class, workflow::cleanup);
    assertFalse(Files.exists(owned));
  }

  private Path closedFilesystemPath() throws IOException {
    try (var filesystem =
        FileSystems.newFileSystem(directory.resolve("closed.zip"), Map.of("create", "true"))) {
      return Files.createDirectory(filesystem.getPath("/owned"));
    }
  }

  private static Set<Path> workflowDirectories(Path temporary) throws IOException {
    try (var paths = Files.list(temporary)) {
      return paths
          .filter(path -> path.getFileName().toString().startsWith("gridgrind-fuzz-workflow-"))
          .collect(Collectors.toSet());
    }
  }
}
