package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Tests for workbook filesystem location records. */
class WorkbookLocationTest {
  @Test
  void storedWorkbookNormalizesPathsAndExposesTheBaseDirectory() {
    WorkbookLocation.StoredWorkbook storedWorkbook =
        new WorkbookLocation.StoredWorkbook(Path.of("tmp", "..", "tmp", "Budget.xlsx"));

    assertEquals(
        Path.of("tmp").toAbsolutePath().normalize(), storedWorkbook.baseDirectory().orElseThrow());
  }

  @Test
  void unsavedWorkbookExposesNoBaseDirectory() {
    assertEquals(
        java.util.Optional.empty(), new WorkbookLocation.UnsavedWorkbook().baseDirectory());
  }
}
