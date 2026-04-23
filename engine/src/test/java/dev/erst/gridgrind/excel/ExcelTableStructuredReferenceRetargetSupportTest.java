package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused regressions for copied-sheet structured-reference retargeting. */
class ExcelTableStructuredReferenceRetargetSupportTest {
  @Test
  void retargetFormulaCellsRewritesStructuredReferencesPrecisely() {
    assertEquals(
        "SUBTOTAL(109,OpsCloseTable_Copy2[Actual])+Table2_Value",
        ExcelTableStructuredReferenceRetargetSupport.retargetFormula(
            "SUBTOTAL(109,Table2[Actual])+Table2_Value",
            java.util.Map.of("Table2", "OpsCloseTable_Copy2")));
  }

  @Test
  void retargetFormulaCellsRejectsBlankTableNamesAndIgnoresIdentityMappings() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellFormula("1+1");

      ExcelTableStructuredReferenceRetargetSupport.retargetFormulaCells(
          sheet, List.of("Sales"), List.of("Sales"));
      assertEquals("1+1", sheet.getRow(0).getCell(0).getCellFormula());

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelTableStructuredReferenceRetargetSupport.retargetFormulaCells(
                      sheet, List.of(" "), List.of("Final")));
      assertEquals("transientTableNames must not contain blank values", failure.getMessage());
    }
  }
}
