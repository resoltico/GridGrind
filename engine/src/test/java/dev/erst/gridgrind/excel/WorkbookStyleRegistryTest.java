package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
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
                  null,
                  null,
                  new ExcelCellFont(null, null, null, null, null, true, null),
                  null,
                  null,
                  null)));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertEquals("Aptos", snapshot.font().fontName());
      assertEquals(300, snapshot.font().fontHeight().twips());
      assertEquals(new BigDecimal("15"), snapshot.font().fontHeight().points());
      assertEquals("#112233", snapshot.font().fontColor());
      assertTrue(snapshot.font().underline());
      assertTrue(snapshot.font().strikeout());
      assertEquals("#FFF2CC", snapshot.fill().foregroundColor());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.border().top().style());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.border().right().style());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.border().bottom().style());
      assertEquals(ExcelBorderStyle.MEDIUM, snapshot.border().left().style());
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
                  new ExcelCellFill(ExcelFillPattern.SOLID, "#DDEBF7", null),
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN),
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE),
                      null,
                      null),
                  null)));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertEquals("#DDEBF7", snapshot.fill().foregroundColor());
      assertEquals(ExcelBorderStyle.THIN, snapshot.border().top().style());
      assertEquals(ExcelBorderStyle.DOUBLE, snapshot.border().right().style());
      assertEquals(ExcelBorderStyle.THIN, snapshot.border().bottom().style());
      assertEquals(ExcelBorderStyle.THIN, snapshot.border().left().style());
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
                  new ExcelCellFont(null, null, "Aptos Display", null, null, false, null),
                  null,
                  new ExcelBorder(
                      null, new ExcelBorderSide(ExcelBorderStyle.THICK), null, null, null),
                  null)));
      ExcelCellStyleSnapshot fontNameSnapshot = styleRegistry.snapshot(cell);
      assertEquals("Aptos Display", fontNameSnapshot.font().fontName());
      assertFalse(fontNameSnapshot.font().underline());
      assertEquals(ExcelBorderStyle.THICK, fontNameSnapshot.border().top().style());
      assertEquals(ExcelBorderStyle.MEDIUM, fontNameSnapshot.border().right().style());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  new ExcelCellFont(
                      null,
                      null,
                      null,
                      ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
                      null,
                      null,
                      null),
                  null,
                  null,
                  null)));
      assertEquals(
          new BigDecimal("11.5"), styleRegistry.snapshot(cell).font().fontHeight().points());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  new ExcelCellFont(null, null, null, null, "#445566", null, null),
                  null,
                  null,
                  null)));
      assertEquals("#445566", styleRegistry.snapshot(cell).font().fontColor());
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
      assertNull(styleRegistry.snapshot(cell).font().fontColor());

      XSSFCellStyle nonRgbFillStyle = workbook.createCellStyle();
      nonRgbFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      nonRgbFillStyle.setFillForegroundColor(new XSSFColor());
      cell.setCellStyle(nonRgbFillStyle);
      assertNull(styleRegistry.snapshot(cell).fill().foregroundColor());

      XSSFCellStyle malformedRgbFillStyle = workbook.createCellStyle();
      malformedRgbFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      malformedRgbFillStyle.setFillForegroundColor(new XSSFColor(new byte[] {0x11, 0x22}));
      cell.setCellStyle(malformedRgbFillStyle);
      assertNull(styleRegistry.snapshot(cell).fill().foregroundColor());
    }
  }

  @Test
  void snapshot_ignoresBorderColorsWhenEffectiveStyleIsNone() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      XSSFCellStyle style = workbook.createCellStyle();
      XSSFColor indexedColor = new XSSFColor();
      indexedColor.setIndexed(IndexedColors.DARK_RED.getIndex());
      style.setBottomBorderColor(indexedColor);
      cell.setCellStyle(style);

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);

      assertEquals(ExcelBorderStyle.NONE, snapshot.border().bottom().style());
      assertNull(snapshot.border().bottom().color());
    }
  }

  @Test
  void mergedStyle_tracksExpandedB4StyleDepth() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  new ExcelCellAlignment(
                      true, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP, 45, 3),
                  new ExcelCellFont(
                      true,
                      false,
                      "Aptos",
                      ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
                      "#102030",
                      true,
                      false),
                  new ExcelCellFill(ExcelFillPattern.THIN_HORIZONTAL_BANDS, "#FFF2CC", "#DDEBF7"),
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN, "#203040"),
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE, "#304050"),
                      null,
                      null),
                  new ExcelCellProtection(false, true))));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertTrue(snapshot.alignment().wrapText());
      assertEquals(45, snapshot.alignment().textRotation());
      assertEquals(3, snapshot.alignment().indentation());
      assertEquals("#102030", snapshot.font().fontColor());
      assertEquals(ExcelFillPattern.THIN_HORIZONTAL_BANDS, snapshot.fill().pattern());
      assertEquals("#FFF2CC", snapshot.fill().foregroundColor());
      assertEquals("#DDEBF7", snapshot.fill().backgroundColor());
      assertEquals("#203040", snapshot.border().top().color());
      assertEquals("#304050", snapshot.border().right().color());
      assertFalse(snapshot.protection().locked());
      assertTrue(snapshot.protection().hiddenFormula());
    }
  }

  @Test
  void defaultStylesAndDateFactoriesPreserveNonFormatState() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      assertEquals(workbook.getCellStyleAt(0).getIndex(), styleRegistry.defaultStyle().getIndex());
      assertEquals("General", styleRegistry.defaultSnapshot().numberFormat());

      cell.setCellStyle(styleRegistry.localDateStyle(cell));
      ExcelCellStyleSnapshot dateSnapshot = styleRegistry.snapshot(cell);
      assertEquals("yyyy-mm-dd", dateSnapshot.numberFormat());
      assertEquals("Aptos", dateSnapshot.font().fontName());
      assertEquals("#FFF2CC", dateSnapshot.fill().foregroundColor());

      cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      ExcelCellStyleSnapshot dateTimeSnapshot = styleRegistry.snapshot(cell);
      assertEquals("yyyy-mm-dd hh:mm:ss", dateTimeSnapshot.numberFormat());
      assertEquals(ExcelBorderStyle.MEDIUM, dateTimeSnapshot.border().top().style());
      assertEquals("#112233", dateTimeSnapshot.font().fontColor());
    }
  }

  @Test
  void mergedStyle_handlesForegroundOnlyFillPatchesAndFillClears() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(cell, fillPatch(new ExcelCellFill(null, "#ABC123", null))));
      assertFillSnapshot(
          styleRegistry.snapshot(cell).fill(), ExcelFillPattern.SOLID, "#ABC123", null);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell, fillPatch(new ExcelCellFill(ExcelFillPattern.BRICKS, "#112233", "#445566"))));
      cell.setCellStyle(
          styleRegistry.mergedStyle(cell, fillPatch(new ExcelCellFill(null, "#AA5500", null))));
      assertFillSnapshot(
          styleRegistry.snapshot(cell).fill(), ExcelFillPattern.BRICKS, "#AA5500", "#445566");

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell, fillPatch(new ExcelCellFill(ExcelFillPattern.BRICKS, null, "#556677"))));
      assertFillSnapshot(
          styleRegistry.snapshot(cell).fill(), ExcelFillPattern.BRICKS, "#AA5500", "#556677");

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell, fillPatch(new ExcelCellFill(ExcelFillPattern.NONE, null, null))));
      assertFillSnapshot(styleRegistry.snapshot(cell).fill(), ExcelFillPattern.NONE, null, null);
    }
  }

  @Test
  void mergedStyle_mergesBorderDefaultsColorsAndExplicitOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN, "#112233"),
                      new ExcelBorderSide(null, "#445566"),
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE, null),
                      null,
                      null))));
      ExcelCellStyleSnapshot mergedBorderSnapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.THIN, mergedBorderSnapshot.border().top().style());
      assertEquals("#445566", mergedBorderSnapshot.border().top().color());
      assertEquals(ExcelBorderStyle.DOUBLE, mergedBorderSnapshot.border().right().style());
      assertEquals("#112233", mergedBorderSnapshot.border().right().color());
      assertEquals("#112233", mergedBorderSnapshot.border().bottom().color());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new ExcelBorder(
                      null, new ExcelBorderSide(ExcelBorderStyle.NONE), null, null, null))));
      ExcelCellStyleSnapshot clearedBorderSnapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.NONE, clearedBorderSnapshot.border().top().style());
      assertNull(clearedBorderSnapshot.border().top().color());
      assertEquals(ExcelBorderStyle.DOUBLE, clearedBorderSnapshot.border().right().style());
      assertEquals("#112233", clearedBorderSnapshot.border().right().color());

      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  cell,
                  borderPatch(
                      new ExcelBorder(
                          null, new ExcelBorderSide(null, "#778899"), null, null, null))));

      Cell blankCell = workbook.getSheet("Budget").createRow(1).createCell(0);
      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  blankCell,
                  borderPatch(
                      new ExcelBorder(
                          null, new ExcelBorderSide(null, "#99AABB"), null, null, null))));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  blankCell,
                  borderPatch(
                      new ExcelBorder(
                          new ExcelBorderSide(ExcelBorderStyle.NONE),
                          new ExcelBorderSide(null, "#CC8844"),
                          null,
                          null,
                          null))));
    }
  }

  @Test
  void mergedStyle_clearsNoneBordersWithoutCrashingWhenNoColorExists() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.NONE), null, null, null, null))));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.NONE, snapshot.border().top().style());
      assertNull(snapshot.border().top().color());
      assertEquals(ExcelBorderStyle.NONE, snapshot.border().right().style());
      assertNull(snapshot.border().right().color());
      assertEquals(ExcelBorderStyle.NONE, snapshot.border().bottom().style());
      assertNull(snapshot.border().bottom().color());
      assertEquals(ExcelBorderStyle.NONE, snapshot.border().left().style());
      assertNull(snapshot.border().left().color());
    }
  }

  @Test
  void mergedStyle_clearsExistingBorderColorsWhenAStyleIsResetToNone() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN, "#112233"),
                      null,
                      null,
                      null,
                      null))));
      assertEquals("#112233", styleRegistry.snapshot(cell).border().top().color());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.NONE), null, null, null, null))));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.NONE, snapshot.border().top().style());
      assertNull(snapshot.border().top().color());
    }
  }

  @Test
  void mergedStyle_appliesProtectionBranchesIndependently() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(cell, protectionPatch(new ExcelCellProtection(false, null))));
      ExcelCellStyleSnapshot lockedSnapshot = styleRegistry.snapshot(cell);
      assertFalse(lockedSnapshot.protection().locked());
      assertFalse(lockedSnapshot.protection().hiddenFormula());

      cell.setCellStyle(
          styleRegistry.mergedStyle(cell, protectionPatch(new ExcelCellProtection(null, true))));
      ExcelCellStyleSnapshot hiddenSnapshot = styleRegistry.snapshot(cell);
      assertFalse(hiddenSnapshot.protection().locked());
      assertTrue(hiddenSnapshot.protection().hiddenFormula());
    }
  }

  @Test
  void snapshot_roundTripsEverySupportedFillPattern() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      var sheet = workbook.createSheet("Patterns");

      ExcelFillPattern[] patterns = ExcelFillPattern.values();
      for (int columnIndex = 0; columnIndex < patterns.length; columnIndex++) {
        ExcelFillPattern pattern = patterns[columnIndex];
        Cell cell = sheet.createRow(0).createCell(columnIndex);
        cell.setCellStyle(styleRegistry.mergedStyle(cell, fillPatch(fillFor(pattern))));
        ExcelCellFillSnapshot snapshot = styleRegistry.snapshot(cell).fill();
        assertFillSnapshot(snapshot, pattern, "#102030", "#405060");
      }
    }
  }

  private ExcelCellStyle fillPatch(ExcelCellFill fill) {
    return new ExcelCellStyle(null, null, null, fill, null, null);
  }

  private ExcelCellStyle borderPatch(ExcelBorder border) {
    return new ExcelCellStyle(null, null, null, null, border, null);
  }

  private ExcelCellStyle protectionPatch(ExcelCellProtection protection) {
    return new ExcelCellStyle(null, null, null, null, null, protection);
  }

  private ExcelCellFill fillFor(ExcelFillPattern pattern) {
    if (pattern == ExcelFillPattern.NONE) {
      return new ExcelCellFill(pattern, null, null);
    }
    if (pattern == ExcelFillPattern.SOLID) {
      return new ExcelCellFill(pattern, "#102030", null);
    }
    return new ExcelCellFill(pattern, "#102030", "#405060");
  }

  private void assertFillSnapshot(
      ExcelCellFillSnapshot snapshot,
      ExcelFillPattern pattern,
      String foregroundColor,
      String backgroundColor) {
    assertEquals(pattern, snapshot.pattern());
    if (pattern == ExcelFillPattern.NONE) {
      assertNull(snapshot.foregroundColor());
      assertNull(snapshot.backgroundColor());
      return;
    }
    assertEquals(foregroundColor, snapshot.foregroundColor());
    if (pattern == ExcelFillPattern.SOLID) {
      assertNull(snapshot.backgroundColor());
      return;
    }
    assertEquals(backgroundColor, snapshot.backgroundColor());
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
