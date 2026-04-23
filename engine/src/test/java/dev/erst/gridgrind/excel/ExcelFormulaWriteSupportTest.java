package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused coverage tests for the centralized formula write support seam. */
class ExcelFormulaWriteSupportTest {
  @Test
  void authoredFormulaWritesWrapInvalidFormulasWithCellLocation() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      InvalidFormulaException exception =
          assertThrows(
              InvalidFormulaException.class,
              () -> ExcelFormulaWriteSupport.setAuthoredFormula(cell, "SUM("));

      assertEquals("Budget", exception.sheetName());
      assertEquals("A1", exception.address());
      assertEquals("SUM(", exception.formula());
      assertEquals("Invalid formula at Budget!A1: SUM(", exception.getMessage());
    }
  }

  @Test
  void rewrittenFormulaWritesSurfaceTheOperationAndCellLocation() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      IllegalStateException exception =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelFormulaWriteSupport.setRewrittenFormula(cell, "SUM(", "copy-sheet rewrite"));

      assertEquals(
          "copy-sheet rewrite produced an invalid formula at Budget!A1: SUM(",
          exception.getMessage());
    }
  }

  @Test
  void scratchFormulaWritesRejectInvalidScratchFormulas() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelFormulaWriteSupport.setScratchFormula(
                      cell, "SUM(", "scratch formula probe"));

      assertTrue(exception.getMessage().contains("scratch formula probe"));
      assertTrue(exception.getMessage().contains("SUM("));
    }
  }
}
