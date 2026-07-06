package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Covers explicit error-value writes through the high-level sheet mutation surface. */
class ExcelSheetErrorWriteCoverageTest {
  @Test
  void sheetCellMutationsPersistOnlyStoredOoxmlErrorLiterals() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");

      sheet.cells().setCell("A1", ExcelCellValue.error("#REF!"));
      sheet.cells().setCell("A2", ExcelCellValue.error("#DIV/0!"));
      sheet.cells().setCell("A3", ExcelCellValue.error("#N/A"));

      ExcelCellSnapshot.ErrorSnapshot snapshot =
          assertInstanceOf(ExcelCellSnapshot.ErrorSnapshot.class, sheet.cells().snapshotCell("A1"));
      assertEquals("#REF!", snapshot.errorValue());
      assertEquals(
          "#DIV/0!",
          assertInstanceOf(ExcelCellSnapshot.ErrorSnapshot.class, sheet.cells().snapshotCell("A2"))
              .errorValue());
      assertEquals(
          "#N/A",
          assertInstanceOf(ExcelCellSnapshot.ErrorSnapshot.class, sheet.cells().snapshotCell("A3"))
              .errorValue());

      Path workbookPath = Files.createTempFile("gridgrind-stored-errors-", ".xlsx");
      workbook
          .persistence()
          .savePlainWorkbook(workbookPath, WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
      try (FileSystem archive = FileSystems.newFileSystem(workbookPath, (ClassLoader) null)) {
        Path sheetXml = archive.getPath("/xl/worksheets/sheet1.xml");
        String xml = Files.readString(sheetXml, StandardCharsets.UTF_8);

        assertTrue(xml.contains("<v>#REF!</v>"));
        assertTrue(xml.contains("<v>#DIV/0!</v>"));
        assertTrue(xml.contains("<v>#N/A</v>"));
        assertFalse(xml.contains("~CIRCULAR~REF~"));
        assertFalse(xml.contains("~FUNCTION~NOT~IMPLEMENTED~"));
      }
    }
  }

  @Test
  void sheetCellMutationsRejectEvaluationOnlyReportedStates() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");

      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setCell("A1", ExcelCellValue.error("#CIRCULAR_REF!")));
      assertThrows(
          IllegalArgumentException.class,
          () -> sheet.cells().setCell("A1", ExcelCellValue.error("#FUNCTION_NOT_IMPLEMENTED!")));
    }
  }
}
