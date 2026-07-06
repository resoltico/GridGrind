package dev.erst.gridgrind.excel.ooxml;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Direct coverage for private OOXML plaintext temp-workbook ownership helpers. */
class ExcelOoxmlPrivateTempWorkbookTest {
  @Test
  void closeDeletesTheOwnedWorkbookTreeWhenOwnershipStaysLocal() throws IOException {
    Path root;
    Path workbookPath;
    try (ExcelOoxmlPrivateTempWorkbook privateWorkbook =
        ExcelOoxmlPrivateTempWorkbook.create("gridgrind-private-", ".xlsx")) {
      root = privateWorkbook.root();
      workbookPath = privateWorkbook.workbookPath();

      assertTrue(Files.isDirectory(root));
      assertEquals(root, workbookPath.getParent());
      assertTrue(Files.exists(workbookPath));
    }

    assertFalse(Files.exists(workbookPath));
    assertFalse(Files.exists(root));
  }

  @Test
  void releaseRootTransfersCleanupOwnershipToTheCaller() throws IOException {
    Path releasedRoot;
    Path workbookPath;
    try (ExcelOoxmlPrivateTempWorkbook privateWorkbook =
        ExcelOoxmlPrivateTempWorkbook.create("gridgrind-private-", ".xlsx")) {
      workbookPath = privateWorkbook.workbookPath();
      releasedRoot = privateWorkbook.releaseRoot();

      assertEquals(privateWorkbook.root(), releasedRoot);
      assertTrue(Files.exists(workbookPath));
    }

    assertTrue(Files.exists(releasedRoot));
    assertTrue(Files.exists(workbookPath));
    ExcelOoxmlPackageFileSupport.deleteTreeIfExists(releasedRoot);
    assertFalse(Files.exists(releasedRoot));
  }
}
