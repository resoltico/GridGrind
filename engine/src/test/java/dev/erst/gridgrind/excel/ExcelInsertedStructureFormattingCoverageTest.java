package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Inserted row and column formatting inheritance coverage. */
class ExcelInsertedStructureFormattingCoverageTest extends ExcelRowColumnStructureTestSupport {
  @Test
  void insertRowsCopiesAdjacentVisualFormattingIntoNewBlankRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      var font = workbook.createFont();
      font.setBold(true);
      font.setFontName("Aptos");
      var style = workbook.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
      style.setFont(font);
      style.setWrapText(true);

      setString(sheet, "A1", "Header");
      Row templateRowOne = sheet.createRow(1);
      templateRowOne.setHeightInPoints(31.5f);
      templateRowOne.setRowStyle(style);
      templateRowOne.createCell(0).setCellValue("North");
      templateRowOne.getCell(0).setCellStyle(style);
      templateRowOne.createCell(1).setCellValue("Approved");
      templateRowOne.getCell(1).setCellStyle(style);

      Row templateRowTwo = sheet.createRow(2);
      templateRowTwo.setHeightInPoints(31.5f);
      templateRowTwo.setRowStyle(style);
      templateRowTwo.createCell(0).setCellValue("South");
      templateRowTwo.getCell(0).setCellStyle(style);
      templateRowTwo.createCell(1).setCellValue("Approved");
      templateRowTwo.getCell(1).setCellStyle(style);

      Row templateRowThree = sheet.createRow(3);
      templateRowThree.setHeightInPoints(31.5f);
      templateRowThree.setRowStyle(style);
      templateRowThree.createCell(0).setCellValue("West");
      templateRowThree.getCell(0).setCellStyle(style);
      templateRowThree.createCell(1).setCellValue("Approved");
      templateRowThree.getCell(1).setCellStyle(style);

      controller.insertRows(sheet, 2, 2);

      for (int rowIndex = 2; rowIndex <= 3; rowIndex++) {
        Row insertedRow = sheet.getRow(rowIndex);
        assertNotNull(insertedRow, "inserted row " + rowIndex + " should be materialized");
        assertEquals(templateRowOne.getHeight(), insertedRow.getHeight());
        assertNotNull(insertedRow.getRowStyle(), "inserted row should carry row-level style");
        assertEquals(templateRowOne.getRowStyle().getIndex(), insertedRow.getRowStyle().getIndex());
        for (int columnIndex = 0; columnIndex <= 1; columnIndex++) {
          Cell insertedCell = insertedRow.getCell(columnIndex);
          assertNotNull(
              insertedCell,
              "inserted row " + rowIndex + " column " + columnIndex + " should be styled");
          assertEquals(
              templateRowOne.getCell(columnIndex).getCellStyle().getIndex(),
              insertedCell.getCellStyle().getIndex());
          assertTrue(
              insertedCell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK
                  || insertedCell.getStringCellValue().isEmpty());
        }
      }
    }
  }

  @Test
  void insertRowsAtTopFallsBackToTheFirstExistingRowBelowForCellFormatting() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      var style = workbook.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
      style.setWrapText(true);

      Row templateRow = sheet.createRow(0);
      templateRow.setHeightInPoints(19.5f);
      templateRow.createCell(0).setCellValue("North");
      templateRow.getCell(0).setCellStyle(style);

      controller.insertRows(sheet, 0, 1);

      Row insertedRow = sheet.getRow(0);
      assertNotNull(insertedRow);
      assertNull(
          insertedRow.getRowStyle(), "cell-only template rows should not fabricate row style");
      assertEquals(templateRow.getHeight(), insertedRow.getHeight());
      assertEquals(
          templateRow.getCell(0).getCellStyle().getIndex(),
          insertedRow.getCell(0).getCellStyle().getIndex());
    }
  }

  @Test
  void insertColumnsCopiesAdjacentVisualFormattingIntoNewBlankColumns() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      var font = workbook.createFont();
      font.setBold(true);
      font.setFontName("Aptos");
      var style = workbook.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
      style.setFont(font);
      style.setWrapText(true);

      sheet.setColumnWidth(0, 4096);
      sheet.setDefaultColumnStyle(0, style);
      setString(sheet, "A1", "North");
      setString(sheet, "A2", "South");
      sheet.getRow(0).getCell(0).setCellStyle(style);
      sheet.getRow(1).getCell(0).setCellStyle(style);
      setString(sheet, "B1", "Shifted");
      setString(sheet, "B2", "Shifted");
      setString(sheet, "B3", "Sparse");

      controller.insertColumns(sheet, 1, 2);

      assertEquals(sheet.getColumnWidth(0), sheet.getColumnWidth(1));
      assertEquals(sheet.getColumnWidth(0), sheet.getColumnWidth(2));
      assertEquals(sheet.getColumnStyle(0).getIndex(), sheet.getColumnStyle(1).getIndex());
      assertEquals(sheet.getColumnStyle(0).getIndex(), sheet.getColumnStyle(2).getIndex());
      for (int rowIndex = 0; rowIndex <= 1; rowIndex++) {
        Row row = sheet.getRow(rowIndex);
        for (int columnIndex = 1; columnIndex <= 2; columnIndex++) {
          Cell insertedCell = row.getCell(columnIndex);
          assertNotNull(insertedCell);
          assertEquals(
              row.getCell(0).getCellStyle().getIndex(), insertedCell.getCellStyle().getIndex());
          assertTrue(
              insertedCell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK
                  || insertedCell.getStringCellValue().isEmpty());
        }
      }
      assertNull(sheet.getRow(2).getCell(1));
      assertNull(sheet.getRow(2).getCell(2));
    }
  }

  @Test
  void insertColumnsAtLeftEdgeFallsBackToTheFirstExistingColumnOnTheRightForVisualFormatting()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      var style = workbook.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
      style.setWrapText(true);

      sheet.setColumnWidth(0, 3584);
      sheet.setDefaultColumnStyle(0, style);
      setString(sheet, "A1", "North");
      sheet.getRow(0).getCell(0).setCellStyle(style);

      controller.insertColumns(sheet, 0, 1);

      assertEquals(3584, sheet.getColumnWidth(0));
      assertEquals(style.getIndex(), sheet.getColumnStyle(0).getIndex());
      assertEquals(style.getIndex(), sheet.getRow(0).getCell(0).getCellStyle().getIndex());
    }
  }

  @Test
  void insertRowsSkipsFormattingMaterializationWhenNoPhysicalTemplateRowExists() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("GhostRows");
      Row ghost = sheet.createRow(0);
      ghost.createCell(0).setCellValue("Ghost");
      sheet.removeRow(ghost);

      controller.insertRows(sheet, 0, 1);

      assertNull(sheet.getRow(0));
    }
  }
}
