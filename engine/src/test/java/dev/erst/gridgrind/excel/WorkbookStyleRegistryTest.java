package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookStyleRegistry utilities, merge behavior, and style snapshots. */
class WorkbookStyleRegistryTest {
  @Test
  void resolveNumberFormat_returnsGeneralForNullOrBlank() {
    assertEquals("General", WorkbookStyleRegistry.resolveNumberFormat(null));
    assertEquals("General", WorkbookStyleRegistry.resolveNumberFormat(""));
    assertEquals("General", WorkbookStyleRegistry.resolveNumberFormat("   "));
  }

  @Test
  void resolveNumberFormat_returnsFormatStringWhenPopulated() {
    assertEquals("#,##0.00", WorkbookStyleRegistry.resolveNumberFormat("#,##0.00"));
    assertEquals("yyyy-mm-dd", WorkbookStyleRegistry.resolveNumberFormat("yyyy-mm-dd"));
  }

  @Test
  void mergedStyle_preservesUnspecifiedFontAttributesFromExistingCellStyle() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null, null, null, null, null, null, null, null, null, true, null, null, null)));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertEquals("Aptos", snapshot.fontName());
      assertEquals(300, snapshot.fontHeight().twips());
      assertEquals(new BigDecimal("15"), snapshot.fontHeight().points());
      assertEquals("#112233", snapshot.fontColor());
      assertTrue(snapshot.underline());
      assertTrue(snapshot.strikeout());
      assertEquals("#FFF2CC", snapshot.fillColor());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.topBorderStyle());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.rightBorderStyle());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.bottomBorderStyle());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.leftBorderStyle());
    }
  }

  @Test
  void mergedStyle_appliesBorderDefaultsAndExplicitSideOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  "#DDEBF7",
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN),
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE),
                      null,
                      null))));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertEquals("#DDEBF7", snapshot.fillColor());
      assertEquals(ExcelBorderStyle.THIN, snapshot.topBorderStyle());
      assertEquals(ExcelBorderStyle.DOUBLE, snapshot.rightBorderStyle());
      assertEquals(ExcelBorderStyle.THIN, snapshot.bottomBorderStyle());
      assertEquals(ExcelBorderStyle.THIN, snapshot.leftBorderStyle());
    }
  }

  @Test
  void mergedStyle_handlesFontPatchVariantsAndPartialBorderOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  "Aptos Display",
                  null,
                  null,
                  false,
                  null,
                  null,
                  new ExcelBorder(
                      null, new ExcelBorderSide(ExcelBorderStyle.THICK), null, null, null))));
      ExcelCellStyleSnapshot fontNameSnapshot = styleRegistry.snapshot(cell);
      assertEquals("Aptos Display", fontNameSnapshot.fontName());
      assertFalse(fontNameSnapshot.underline());
      assertEquals(ExcelBorderStyle.THICK, fontNameSnapshot.topBorderStyle());
      assertEquals(ExcelBorderStyle.MEDIUM, fontNameSnapshot.rightBorderStyle());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  null,
                  ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
                  null,
                  null,
                  null,
                  null,
                  null)));
      assertEquals(new BigDecimal("11.5"), styleRegistry.snapshot(cell).fontHeight().points());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null, null, null, null, null, null, null, null, "#445566", null, null, null,
                  null)));
      assertEquals("#445566", styleRegistry.snapshot(cell).fontColor());
    }
  }

  @Test
  void snapshot_handlesNullAndNonRgbColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      XSSFCellStyle noColorFontStyle = workbook.createCellStyle();
      XSSFFont noColorFont = workbook.createFont();
      noColorFontStyle.setFont(noColorFont);
      cell.setCellStyle(noColorFontStyle);
      assertNull(styleRegistry.snapshot(cell).fontColor());

      XSSFCellStyle nonRgbFillStyle = workbook.createCellStyle();
      nonRgbFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      nonRgbFillStyle.setFillForegroundColor(new XSSFColor());
      cell.setCellStyle(nonRgbFillStyle);
      assertNull(styleRegistry.snapshot(cell).fillColor());

      XSSFCellStyle malformedRgbFillStyle = workbook.createCellStyle();
      malformedRgbFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      malformedRgbFillStyle.setFillForegroundColor(new XSSFColor(new byte[] {0x11, 0x22}));
      cell.setCellStyle(malformedRgbFillStyle);
      assertNull(styleRegistry.snapshot(cell).fillColor());
    }
  }

  private XSSFCellStyle styledBaseCellStyle(XSSFWorkbook workbook) {
    XSSFCellStyle style = workbook.createCellStyle();
    style.setFillForegroundColor(new XSSFColor(new byte[] {(byte) 0xFF, (byte) 0xF2, (byte) 0xCC}));
    style.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
    style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
    style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
    style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
    style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);

    XSSFFont font = workbook.createFont();
    font.setFontName("Aptos");
    font.setFontHeightInPoints((short) 15);
    font.setColor(new XSSFColor(new byte[] {0x11, 0x22, 0x33}));
    font.setStrikeout(true);
    style.setFont(font);
    return style;
  }
}
