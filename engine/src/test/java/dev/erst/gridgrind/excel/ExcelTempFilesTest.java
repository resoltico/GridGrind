package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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
  void managedTempFilesUseExplicitRootsWhenSupplied() throws IOException {
    Path explicitRoot = Files.createTempDirectory("gridgrind-explicit-root-");
    Path tempFile = null;
    Path tempDirectory = null;
    try {
      tempFile = ExcelTempFiles.createManagedTempFile(explicitRoot, "gridgrind-explicit-", ".xlsx");
      tempDirectory =
          ExcelTempFiles.createManagedTempDirectory(explicitRoot, "gridgrind-explicit-dir-");
      assertTrue(tempFile.startsWith(explicitRoot.toAbsolutePath().normalize()));
      assertTrue(tempDirectory.startsWith(explicitRoot.toAbsolutePath().normalize()));
    } finally {
      Files.deleteIfExists(tempFile);
      Files.deleteIfExists(tempDirectory);
      Files.deleteIfExists(explicitRoot);
    }
  }

  @Test
  void managedTempFilesRejectBrokenSystemTempRootsInsteadOfFallingBackToUserHome()
      throws IOException {
    Path invalidSystemTemp = Files.createTempFile("gridgrind-temp-root-file-", ".tmp");
    String originalSystemTemp = System.getProperty("java.io.tmpdir");
    System.setProperty("java.io.tmpdir", invalidSystemTemp.toString());
    try {
      assertThrows(
          IOException.class,
          () -> ExcelTempFiles.createManagedTempFile("gridgrind-fallback-", ".xlsx"));
      assertThrows(
          IOException.class,
          () -> ExcelTempFiles.createManagedTempDirectory("gridgrind-fallback-dir-"));
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
      Files.deleteIfExists(invalidSystemTemp);
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
    System.setProperty("java.io.tmpdir", "");
    try {
      assertThrows(
          IOException.class,
          () -> ExcelTempFiles.createManagedTempFile("gridgrind-none-", ".xlsx"));
      assertThrows(
          IOException.class,
          () -> ExcelTempFiles.createManagedTempDirectory("gridgrind-none-dir-"));
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
    }
  }

  @Test
  void managedTempFilesPropagateSystemTempFailures() throws IOException {
    Path invalidSystemTemp = Files.createTempFile("gridgrind-bad-system-root-", ".tmp");
    String originalSystemTemp = System.getProperty("java.io.tmpdir");
    System.setProperty("java.io.tmpdir", invalidSystemTemp.toString());
    try {
      IOException fileFailure =
          assertThrows(
              IOException.class,
              () -> ExcelTempFiles.createManagedTempFile("gridgrind-fail-", ".tmp"));
      assertEquals(0, fileFailure.getSuppressed().length);

      IOException directoryFailure =
          assertThrows(
              IOException.class,
              () -> ExcelTempFiles.createManagedTempDirectory("gridgrind-fail-dir-"));
      assertEquals(0, directoryFailure.getSuppressed().length);
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
      Files.deleteIfExists(invalidSystemTemp);
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
