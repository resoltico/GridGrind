package dev.erst.gridgrind.engine.security;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.engine.packaging.WorkbookFiles;
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
    try (PrivateWorkbookFile privateWorkbook =
        PrivateWorkbookFile.create("gridgrind-private-", ".xlsx")) {
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
    try (PrivateWorkbookFile privateWorkbook =
        PrivateWorkbookFile.create("gridgrind-private-", ".xlsx")) {
      workbookPath = privateWorkbook.workbookPath();
      releasedRoot = privateWorkbook.releaseRoot();

      assertEquals(privateWorkbook.root(), releasedRoot);
      assertTrue(Files.exists(workbookPath));
    }

    assertTrue(Files.exists(releasedRoot));
    assertTrue(Files.exists(workbookPath));
    WorkbookFiles.deleteTreeIfExists(releasedRoot);
    assertFalse(Files.exists(releasedRoot));
  }
}
