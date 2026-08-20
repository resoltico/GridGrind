package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
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
  void authoredFormulaRejectsPoiUnparseableSyntaxButAcceptsUnknownFunctionNames()
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFCell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      ExcelFormulaWriteSupport.setAuthoredFormula(cell, "ZZZZ(1)");
      assertEquals("ZZZZ(1)", cell.getCellFormula());
      UnsupportedFormulaConstructException exception =
          assertThrows(
              UnsupportedFormulaConstructException.class,
              () -> ExcelFormulaWriteSupport.setAuthoredFormula(cell, "LAMBDA(x,x+1)(A1)"));

      assertEquals("LAMBDA(x,x+1)(A1)", exception.formula());
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

  @Test
  void opaqueFormulaWritesPersistFormulaCharacterDataWithoutPoiParsing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFCell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      ExcelFormulaWriteSupport.setOpaqueFormula(cell, "LAMBDA(x,x+1)(A1)");

      assertEquals("LAMBDA(x,x+1)(A1)", cell.getCTCell().getF().getStringValue());
      assertFalse(cell.getCTCell().isSetV());
    }
  }

  @Test
  void opaqueFormulaReplacesAnExistingPoiFormulaAndCachedValue() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFCell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellFormula("1+1");
      cell.setCellValue(2.0d);

      ExcelFormulaWriteSupport.setOpaqueFormula(cell, "LET(x,1,x+1)");

      assertEquals("LET(x,1,x+1)", cell.getCTCell().getF().getStringValue());
      assertFalse(cell.getCTCell().isSetV());
      assertFalse(cell.getCTCell().isSetT());
    }
  }

  @Test
  void opaqueFormulaRejectsXmlForbiddenControlsButKeepsMarkupLookingText() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFCell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      ExcelFormulaWriteSupport.setOpaqueFormula(cell, "\"<b>\"&A1");
      assertEquals("\"<b>\"&A1", cell.getCTCell().getF().getStringValue());
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelFormulaWriteSupport.setOpaqueFormula(cell, "A1\u0001+B1"));
    }
  }

  @Test
  void opaqueFormulaWritesEscapedCharacterDataThatSurvivesAnOoxmlRoundTrip() throws IOException {
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream output = new ByteArrayOutputStream()) {
      XSSFCell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      ExcelFormulaWriteSupport.setOpaqueFormula(cell, "\"<b>\"&A1");
      XSSFCell lambdaCell = cell.getRow().createCell(1);
      ExcelFormulaWriteSupport.setOpaqueFormula(lambdaCell, "LAMBDA(x,x+1)(A1)");

      workbook.write(output);
      bytes = output.toByteArray();
    }

    String worksheetXml = worksheetXml(bytes);
    assertTrue(worksheetXml.contains("&lt;b>"));
    assertTrue(worksheetXml.contains("&amp;A1"));
    assertFalse(worksheetXml.contains("\"<b>\"&A1"));
    assertTrue(worksheetXml.contains("LAMBDA(x,x+1)(A1)"));
    try (XSSFWorkbook reopened = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
      assertEquals(
          "\"<b>\"&A1",
          reopened.getSheet("Budget").getRow(0).getCell(0).getCTCell().getF().getStringValue());
      assertEquals(
          "LAMBDA(x,x+1)(A1)",
          reopened.getSheet("Budget").getRow(0).getCell(1).getCTCell().getF().getStringValue());
    }
  }

  @Test
  void opaqueFormulaRequiresAnXssfCell() throws IOException {
    try (HSSFWorkbook workbook = new HSSFWorkbook()) {
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      assertEquals(
          "Opaque formulas require an XSSF or SXSSF cell",
          assertThrows(
                  IllegalArgumentException.class,
                  () -> ExcelFormulaWriteSupport.setOpaqueFormula(cell, "LAMBDA(x,x)(A1)"))
              .getMessage());
    }
  }

  private static String worksheetXml(byte[] workbookBytes) throws IOException {
    try (ZipInputStream zip = new ZipInputStream(new ByteArrayInputStream(workbookBytes))) {
      for (ZipEntry entry = zip.getNextEntry(); entry != null; entry = zip.getNextEntry()) {
        if ("xl/worksheets/sheet1.xml".equals(entry.getName())) {
          return new String(zip.readAllBytes(), StandardCharsets.UTF_8);
        }
      }
    }
    throw new IllegalStateException("worksheet XML is missing");
  }
}
