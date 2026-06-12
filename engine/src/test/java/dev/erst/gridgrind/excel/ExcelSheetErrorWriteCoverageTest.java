package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import org.junit.jupiter.api.Test;

/** Covers explicit error-value writes through the high-level sheet mutation surface. */
class ExcelSheetErrorWriteCoverageTest {
  @Test
  void sheetCellMutationsPersistExplicitExcelErrorValues() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");

      sheet.cells().setCell("A1", ExcelCellValue.error("#REF!"));

      ExcelCellSnapshot.ErrorSnapshot snapshot =
          assertInstanceOf(ExcelCellSnapshot.ErrorSnapshot.class, sheet.cells().snapshotCell("A1"));
      assertEquals("#REF!", snapshot.errorValue());
    }
  }
}
