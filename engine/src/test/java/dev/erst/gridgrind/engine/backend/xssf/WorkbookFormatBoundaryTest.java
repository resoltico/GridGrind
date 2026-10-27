package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies the content-format boundary independently of the filename suffix. */
class WorkbookFormatBoundaryTest {
  @Test
  void binaryXlsContentIsRejectedEvenWithAnXlsxName(@TempDir Path root) throws Exception {
    Path renamed = root.resolve("renamed.xlsx");
    try (var workbook = new HSSFWorkbook();
        var output = Files.newOutputStream(renamed)) {
      workbook.createSheet("Binary");
      workbook.write(output);
    }
    assertThrows(
        WorkbookNotOpenableException.class,
        () -> XssfWorkbooks.open(renamed, WorkbookTempFileFactory.rooted(root.resolve("scratch"))));
    Path valid = root.resolve("valid.xlsx");
    try (var workbook = XssfWorkbooks.create()) {
      workbook.sheets().getOrCreateSheet("OOXML");
      workbook
          .persistence()
          .save(
              valid,
              WorkbookArtifactWriteDisposition.CREATE_NEW,
              WorkbookTempFileFactory.rooted(root.resolve("write-scratch")));
    }
    try (var reopened =
        XssfWorkbooks.open(valid, WorkbookTempFileFactory.rooted(root.resolve("read-scratch")))) {
      assertEquals(1, reopened.sheets().sheetCount());
    }
  }
}
