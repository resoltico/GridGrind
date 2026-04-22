package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Regression coverage for the package-owned managed temp-file factory. */
class ExcelTempFilesTest {
  @Test
  void managedTempFilesUseDedicatedGridGrindRoots() throws IOException {
    Path tempFile = ExcelTempFiles.createManagedTempFile("gridgrind-managed-", ".xlsx");
    Path tempDirectory = ExcelTempFiles.createManagedTempDirectory("gridgrind-managed-dir-");
    try {
      assertTrue(Files.exists(tempFile));
      assertTrue(tempFile.getFileName().toString().startsWith("gridgrind-managed-"));
      assertTrue(tempFile.getFileName().toString().endsWith(".xlsx"));
      assertEquals("gridgrind", tempFile.getParent().getFileName().toString());
      assertTrue(Files.isDirectory(tempDirectory));
      assertEquals("gridgrind", tempDirectory.getParent().getFileName().toString());
    } finally {
      Files.deleteIfExists(tempFile);
      Files.deleteIfExists(tempDirectory);
    }
  }

  @Test
  void managedTempFilesFallBackFromBrokenSystemTempRootsToUserHomeRoots() throws IOException {
    Path invalidSystemTemp = Files.createTempFile("gridgrind-temp-root-file-", ".tmp");
    Path fallbackUserHome = Files.createTempDirectory("gridgrind-temp-home-");
    String originalSystemTemp = System.getProperty("java.io.tmpdir");
    String originalUserHome = System.getProperty("user.home");
    System.setProperty("java.io.tmpdir", invalidSystemTemp.toString());
    System.setProperty("user.home", fallbackUserHome.toString());

    Path tempFile = null;
    Path tempDirectory = null;
    try {
      tempFile = ExcelTempFiles.createManagedTempFile("gridgrind-fallback-", ".xlsx");
      tempDirectory = ExcelTempFiles.createManagedTempDirectory("gridgrind-fallback-dir-");
      assertTrue(tempFile.startsWith(fallbackUserHome.resolve(".gridgrind").resolve("tmp")));
      assertTrue(tempDirectory.startsWith(fallbackUserHome.resolve(".gridgrind").resolve("tmp")));
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
      restoreProperty("user.home", originalUserHome);
      Files.deleteIfExists(tempFile);
      Files.deleteIfExists(tempDirectory);
      Files.deleteIfExists(invalidSystemTemp);
      if (Files.exists(fallbackUserHome.resolve(".gridgrind").resolve("tmp"))) {
        try (Stream<Path> paths = Files.walk(fallbackUserHome.resolve(".gridgrind"))) {
          paths
              .sorted(Comparator.reverseOrder())
              .forEach(
                  path -> {
                    try {
                      Files.deleteIfExists(path);
                    } catch (IOException ignored) {
                      // Best-effort cleanup for test-owned temp roots.
                    }
                  });
        }
      }
      Files.deleteIfExists(fallbackUserHome);
    }
  }

  @Test
  void managedTempFilesRejectNullsAndFailWhenNoCandidateRootExists() {
    assertThrows(
        NullPointerException.class, () -> ExcelTempFiles.createManagedTempFile(null, ".xlsx"));
    assertThrows(
        NullPointerException.class, () -> ExcelTempFiles.createManagedTempFile("prefix-", null));
    assertThrows(NullPointerException.class, () -> ExcelTempFiles.createManagedTempDirectory(null));

    String originalSystemTemp = System.getProperty("java.io.tmpdir");
    String originalUserHome = System.getProperty("user.home");
    System.setProperty("java.io.tmpdir", "");
    System.setProperty("user.home", "");
    try {
      assertThrows(
          IOException.class,
          () -> ExcelTempFiles.createManagedTempFile("gridgrind-none-", ".xlsx"));
      assertThrows(
          IOException.class,
          () -> ExcelTempFiles.createManagedTempDirectory("gridgrind-none-dir-"));
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
      restoreProperty("user.home", originalUserHome);
    }
  }

  @Test
  void managedTempFilesPropagateThePrimaryFailureAfterAllCandidateRootsFail() throws IOException {
    Path invalidSystemTemp = Files.createTempFile("gridgrind-bad-system-root-", ".tmp");
    Path invalidUserHome = Files.createTempFile("gridgrind-bad-home-root-", ".tmp");
    String originalSystemTemp = System.getProperty("java.io.tmpdir");
    String originalUserHome = System.getProperty("user.home");
    System.setProperty("java.io.tmpdir", invalidSystemTemp.toString());
    System.setProperty("user.home", invalidUserHome.toString());
    try {
      IOException fileFailure =
          assertThrows(
              IOException.class,
              () -> ExcelTempFiles.createManagedTempFile("gridgrind-fail-", ".tmp"));
      assertTrue(fileFailure.getSuppressed().length >= 1);

      IOException directoryFailure =
          assertThrows(
              IOException.class,
              () -> ExcelTempFiles.createManagedTempDirectory("gridgrind-fail-dir-"));
      assertTrue(directoryFailure.getSuppressed().length >= 1);
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
      restoreProperty("user.home", originalUserHome);
      Files.deleteIfExists(invalidSystemTemp);
      Files.deleteIfExists(invalidUserHome);
    }
  }

  private static void restoreProperty(String name, String value) {
    if (value == null) {
      System.clearProperty(name);
    } else {
      System.setProperty(name, value);
    }
  }
}
