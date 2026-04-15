package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
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
      assertEquals(rgb("#112233"), snapshot.font().fontColor());
      assertTrue(snapshot.font().underline());
      assertTrue(snapshot.font().strikeout());
      assertEquals(rgb("#FFF2CC"), snapshot.fill().foregroundColor());
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
      assertEquals(rgb("#DDEBF7"), snapshot.fill().foregroundColor());
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
                  new ExcelCellFont(null, null, null, null, new ExcelColor("#445566"), null, null),
                  null,
                  null,
                  null)));
      assertEquals(rgb("#445566"), styleRegistry.snapshot(cell).font().fontColor());
    }
  }

  @Test
  void mergedStyle_preservesThemeIndexedAndTintFontColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  new ExcelCellFont(
                      null, true, null, null, new ExcelColor(null, 6, null, -0.35d), null, null),
                  null,
                  null,
                  null)));
      assertEquals(
          new ExcelColorSnapshot(null, 6, null, -0.35d),
          styleRegistry.snapshot(cell).font().fontColor());

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
                      null,
                      new ExcelColor(
                          null, null, Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()), null),
                      null,
                      null),
                  null,
                  null,
                  null)));
      assertEquals(
          new ExcelColorSnapshot(
              null, null, Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()), null),
          styleRegistry.snapshot(cell).font().fontColor());
    }
  }

  @Test
  void mergedStyle_appliesGradientFillPatchesAndCollapsesExtraFontColorNodes() throws Exception {
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
                  new ExcelCellFont(
                      null, null, null, null, new ExcelColor(null, 6, null, -0.35d), null, null),
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "PATH",
                          null,
                          0.1d,
                          0.2d,
                          0.3d,
                          0.4d,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                              new ExcelGradientStop(1.0d, new ExcelColor(null, null, 5, null))))),
                  null,
                  null)));

      XSSFCellStyle snapshotStyle = (XSSFCellStyle) cell.getCellStyle();
      XSSFFont snapshotFont = workbook.getFontAt(snapshotStyle.getFontIndexAsInt());
      var gradientFill =
          workbook
              .getStylesSource()
              .getFillAt((int) snapshotStyle.getCoreXf().getFillId())
              .getCTFill()
              .getGradientFill();

      assertEquals(6L, snapshotFont.getCTFont().getColorArray(0).getTheme());
      assertEquals(-0.35d, snapshotFont.getCTFont().getColorArray(0).getTint());
      assertNotNull(gradientFill);
      assertEquals("path", gradientFill.getType().toString());
      assertEquals(0.1d, gradientFill.getLeft());
      assertEquals(
          "#112233",
          ExcelColorSnapshotSupport.snapshot(workbook, gradientFill.getStopArray(0).getColor())
              .rgb());
      assertEquals(
          5,
          ExcelColorSnapshotSupport.snapshot(workbook, gradientFill.getStopArray(1).getColor())
              .indexed());
    }
  }

  @Test
  void mergedStyle_collapsesDuplicateFontColorsAndSupportsLinearGradientDefaults()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      XSSFCellStyle style = workbook.createCellStyle();
      XSSFFont font = workbook.createFont();
      font.setFontName("Aptos");
      font.getCTFont().addNewColor().setRgb(new byte[] {0x01, 0x02, 0x03});
      font.getCTFont().addNewColor().setTheme(2L);
      font.getCTFont().addNewColor().setIndexed(5L);
      style.setFont(font);
      cell.setCellStyle(style);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  new ExcelCellFont(null, null, null, null, new ExcelColor("#445566"), null, null),
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "LINEAR",
                          37.5d,
                          null,
                          null,
                          null,
                          null,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                              new ExcelGradientStop(1.0d, new ExcelColor("#445566"))))),
                  null,
                  null)));

      XSSFCellStyle mergedStyle = (XSSFCellStyle) cell.getCellStyle();
      XSSFFont mergedFont = workbook.getFontAt(mergedStyle.getFontIndexAsInt());
      var gradientFill =
          workbook
              .getStylesSource()
              .getFillAt((int) mergedStyle.getCoreXf().getFillId())
              .getCTFill()
              .getGradientFill();

      assertEquals(1, mergedFont.getCTFont().sizeOfColorArray());
      assertEquals(
          "#445566",
          ExcelColorSnapshotSupport.snapshot(workbook, mergedFont.getCTFont().getColorArray(0))
              .rgb());
      assertNotNull(gradientFill);
      assertFalse(gradientFill.isSetType());
      assertEquals(37.5d, gradientFill.getDegree());
    }
  }

  @Test
  void mergedStyle_appliesFontColorWhenBaseFontHasNoColorNode() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      XSSFCellStyle style = workbook.createCellStyle();
      XSSFFont font = workbook.createFont();
      font.setFontName("Aptos");
      while (font.getCTFont().sizeOfColorArray() > 0) {
        font.getCTFont().removeColor(font.getCTFont().sizeOfColorArray() - 1);
      }
      style.setFont(font);
      cell.setCellStyle(style);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  new ExcelCellFont(null, null, null, null, new ExcelColor("#102030"), null, null),
                  null,
                  null,
                  null)));

      XSSFCellStyle mergedStyle = (XSSFCellStyle) cell.getCellStyle();
      XSSFFont mergedFont = workbook.getFontAt(mergedStyle.getFontIndexAsInt());
      assertEquals(1, mergedFont.getCTFont().sizeOfColorArray());
      assertEquals(
          "#102030",
          ExcelColorSnapshotSupport.snapshot(workbook, mergedFont.getCTFont().getColorArray(0))
              .rgb());
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
  void snapshot_preservesBorderColorsWhenEffectiveStyleIsNone() throws Exception {
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
      assertEquals(
          new ExcelColorSnapshot(null, null, (int) IndexedColors.DARK_RED.getIndex(), null),
          snapshot.border().bottom().color());
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
                      new ExcelColor("#102030"),
                      true,
                      false),
                  new ExcelCellFill(ExcelFillPattern.THIN_HORIZONTAL_BANDS, "#FFF2CC", "#DDEBF7"),
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN, new ExcelColor("#203040")),
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE, new ExcelColor("#304050")),
                      null,
                      null),
                  new ExcelCellProtection(false, true))));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);
      assertTrue(snapshot.alignment().wrapText());
      assertEquals(45, snapshot.alignment().textRotation());
      assertEquals(3, snapshot.alignment().indentation());
      assertEquals(rgb("#102030"), snapshot.font().fontColor());
      assertEquals(ExcelFillPattern.THIN_HORIZONTAL_BANDS, snapshot.fill().pattern());
      assertEquals(rgb("#FFF2CC"), snapshot.fill().foregroundColor());
      assertEquals(rgb("#DDEBF7"), snapshot.fill().backgroundColor());
      assertEquals(rgb("#203040"), snapshot.border().top().color());
      assertEquals(rgb("#304050"), snapshot.border().right().color());
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
      assertEquals(rgb("#FFF2CC"), dateSnapshot.fill().foregroundColor());

      cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      ExcelCellStyleSnapshot dateTimeSnapshot = styleRegistry.snapshot(cell);
      assertEquals("yyyy-mm-dd hh:mm:ss", dateTimeSnapshot.numberFormat());
      assertEquals(ExcelBorderStyle.MEDIUM, dateTimeSnapshot.border().top().style());
      assertEquals(rgb("#112233"), dateTimeSnapshot.font().fontColor());
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
                      new ExcelBorderSide(ExcelBorderStyle.THIN, new ExcelColor("#112233")),
                      new ExcelBorderSide(null, new ExcelColor("#445566")),
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE, null),
                      null,
                      null))));
      ExcelCellStyleSnapshot mergedBorderSnapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.THIN, mergedBorderSnapshot.border().top().style());
      assertEquals(rgb("#445566"), mergedBorderSnapshot.border().top().color());
      assertEquals(ExcelBorderStyle.DOUBLE, mergedBorderSnapshot.border().right().style());
      assertEquals(rgb("#112233"), mergedBorderSnapshot.border().right().color());
      assertEquals(rgb("#112233"), mergedBorderSnapshot.border().bottom().color());

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
      assertEquals(rgb("#112233"), clearedBorderSnapshot.border().right().color());

      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  cell,
                  borderPatch(
                      new ExcelBorder(
                          null,
                          new ExcelBorderSide(null, new ExcelColor("#778899")),
                          null,
                          null,
                          null))));

      Cell blankCell = workbook.getSheet("Budget").createRow(1).createCell(0);
      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  blankCell,
                  borderPatch(
                      new ExcelBorder(
                          null,
                          new ExcelBorderSide(null, new ExcelColor("#99AABB")),
                          null,
                          null,
                          null))));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  blankCell,
                  borderPatch(
                      new ExcelBorder(
                          new ExcelBorderSide(ExcelBorderStyle.NONE),
                          new ExcelBorderSide(null, new ExcelColor("#CC8844")),
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
                      new ExcelBorderSide(ExcelBorderStyle.THIN, new ExcelColor("#112233")),
                      null,
                      null,
                      null,
                      null))));
      assertEquals(rgb("#112233"), styleRegistry.snapshot(cell).border().top().color());

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

  @Test
  void snapshot_readsGradientFillStopsAndWorkbookColorSemantics() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      XSSFCellStyle style = workbook.createCellStyle();
      XSSFCellFill styleFill =
          workbook.getStylesSource().getFillAt((int) style.getCoreXf().getFillId());
      if (styleFill.getCTFill().isSetPatternFill()) {
        styleFill.getCTFill().unsetPatternFill();
      }
      var gradientFill = styleFill.getCTFill().addNewGradientFill();
      gradientFill.setDegree(45.0d);
      var firstStop = gradientFill.addNewStop();
      firstStop.setPosition(0.0d);
      firstStop.addNewColor().setRgb(new byte[] {(byte) 0xFF, 0x11, 0x22, 0x33});
      var secondStop = gradientFill.addNewStop();
      secondStop.setPosition(1.0d);
      secondStop.addNewColor().setIndexed(10L);
      cell.setCellStyle(style);

      ExcelCellFillSnapshot snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.NONE, snapshot.pattern());
      assertNotNull(snapshot.gradient());
      assertEquals("LINEAR", snapshot.gradient().type());
      assertEquals(45.0d, snapshot.gradient().degree());
      assertEquals(2, snapshot.gradient().stops().size());
      assertEquals(new ExcelColorSnapshot("#112233"), snapshot.gradient().stops().get(0).color());
      assertEquals(
          new ExcelColorSnapshot(null, null, 10, null), snapshot.gradient().stops().get(1).color());
    }
  }

  @Test
  void gradientFillSnapshotReflectsExplicitPathOffsets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      var gradient =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill.Factory.newInstance();
      gradient.setType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STGradientType.PATH);
      gradient.setLeft(0.1d);
      gradient.setRight(0.2d);
      gradient.setTop(0.3d);
      gradient.setBottom(0.4d);
      var firstStop = gradient.addNewStop();
      firstStop.setPosition(0.0d);
      firstStop.addNewColor().setTheme(2L);
      var secondStop = gradient.addNewStop();
      secondStop.setPosition(1.0d);
      secondStop.addNewColor().setIndexed(5L);

      ExcelGradientFillSnapshot snapshot = styleRegistry.gradientFillSnapshot(gradient);

      assertEquals("PATH", snapshot.type());
      assertNull(snapshot.degree());
      assertEquals(0.1d, snapshot.left());
      assertEquals(0.2d, snapshot.right());
      assertEquals(0.3d, snapshot.top());
      assertEquals(0.4d, snapshot.bottom());
      assertEquals(new ExcelColorSnapshot(null, 2, null, null), snapshot.stops().get(0).color());
      assertEquals(new ExcelColorSnapshot(null, null, 5, null), snapshot.stops().get(1).color());
    }
  }

  @Test
  void gradientFillSnapshotInfersPathTypeWhenOffsetsExistWithoutExplicitType() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      var gradient =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill.Factory.newInstance();
      gradient.setLeft(0.1d);
      gradient.setRight(0.2d);
      gradient.setTop(0.3d);
      gradient.setBottom(0.4d);
      var firstStop = gradient.addNewStop();
      firstStop.setPosition(0.0d);
      firstStop.addNewColor().setTheme(2L);
      var secondStop = gradient.addNewStop();
      secondStop.setPosition(1.0d);
      secondStop.addNewColor().setIndexed(5L);

      ExcelGradientFillSnapshot snapshot = styleRegistry.gradientFillSnapshot(gradient);

      assertEquals("PATH", snapshot.type());
      assertNull(snapshot.degree());
      assertEquals(0.1d, snapshot.left());
      assertEquals(0.2d, snapshot.right());
      assertEquals(0.3d, snapshot.top());
      assertEquals(0.4d, snapshot.bottom());
    }
  }

  @Test
  void snapshotReportsPathGradientGeometryForStyledCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellValue("Path");

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "PATH",
                          null,
                          0.1d,
                          0.2d,
                          0.3d,
                          0.4d,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                              new ExcelGradientStop(1.0d, new ExcelColor(null, null, 10, null))))),
                  null,
                  new ExcelCellProtection(false, true))));

      ExcelCellStyleSnapshot snapshot = styleRegistry.snapshot(cell);

      assertEquals("PATH", snapshot.fill().gradient().type());
      assertNull(snapshot.fill().gradient().degree());
      assertEquals(0.1d, snapshot.fill().gradient().left());
      assertEquals(0.2d, snapshot.fill().gradient().right());
      assertEquals(0.3d, snapshot.fill().gradient().top());
      assertEquals(0.4d, snapshot.fill().gradient().bottom());
      assertFalse(snapshot.protection().locked());
      assertTrue(snapshot.protection().hiddenFormula());
    }
  }

  @Test
  void distinctGradientFillsDoNotAliasWhenWorkbookUsesMultipleGradientStyles() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      var sheet = workbook.createSheet("Budget");
      Cell linearCell = sheet.createRow(0).createCell(0);
      Cell pathCell = sheet.getRow(0).createCell(1);
      linearCell.setCellValue("Linear");
      pathCell.setCellValue("Path");

      linearCell.setCellStyle(
          styleRegistry.mergedStyle(
              linearCell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "LINEAR",
                          45.0d,
                          null,
                          null,
                          null,
                          null,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#1F497D")),
                              new ExcelGradientStop(1.0d, new ExcelColor(null, 4, null, 0.45d))))),
                  null,
                  new ExcelCellProtection(true, true))));
      pathCell.setCellStyle(
          styleRegistry.mergedStyle(
              pathCell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "PATH",
                          null,
                          0.1d,
                          0.2d,
                          0.3d,
                          0.4d,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                              new ExcelGradientStop(1.0d, new ExcelColor(null, null, 10, null))))),
                  null,
                  new ExcelCellProtection(false, true))));

      ExcelCellStyleSnapshot linearSnapshot = styleRegistry.snapshot(linearCell);
      ExcelCellStyleSnapshot pathSnapshot = styleRegistry.snapshot(pathCell);

      assertEquals("LINEAR", linearSnapshot.fill().gradient().type());
      assertEquals(45.0d, linearSnapshot.fill().gradient().degree());
      assertEquals("PATH", pathSnapshot.fill().gradient().type());
      assertEquals(0.1d, pathSnapshot.fill().gradient().left());
      assertEquals(0.2d, pathSnapshot.fill().gradient().right());
      assertEquals(0.3d, pathSnapshot.fill().gradient().top());
      assertEquals(0.4d, pathSnapshot.fill().gradient().bottom());
      assertFalse(pathSnapshot.protection().locked());
      assertTrue(pathSnapshot.protection().hiddenFormula());
    }
  }

  @Test
  void identicalGradientFillsReuseOneStylesTableFillId() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      var sheet = workbook.createSheet("Budget");
      Cell firstCell = sheet.createRow(0).createCell(0);
      Cell secondCell = sheet.getRow(0).createCell(1);
      firstCell.setCellValue("Locked");
      secondCell.setCellValue("Unlocked");
      ExcelGradientFill sharedGradient =
          new ExcelGradientFill(
              "LINEAR",
              45.0d,
              null,
              null,
              null,
              null,
              List.of(
                  new ExcelGradientStop(0.0d, new ExcelColor("#1F497D")),
                  new ExcelGradientStop(1.0d, new ExcelColor(null, 4, null, 0.45d))));

      firstCell.setCellStyle(
          styleRegistry.mergedStyle(
              firstCell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(null, null, null, sharedGradient),
                  null,
                  new ExcelCellProtection(true, true))));
      secondCell.setCellStyle(
          styleRegistry.mergedStyle(
              secondCell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(null, null, null, sharedGradient),
                  null,
                  new ExcelCellProtection(false, true))));

      long firstFillId = ((XSSFCellStyle) firstCell.getCellStyle()).getCoreXf().getFillId();
      long secondFillId = ((XSSFCellStyle) secondCell.getCellStyle()).getCoreXf().getFillId();

      assertEquals(firstFillId, secondFillId);
    }
  }

  @Test
  void pathGradientsWriteOnlyProvidedOffsets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      WorkbookStyleRegistry styleRegistry = new WorkbookStyleRegistry(workbook);
      var sheet = workbook.createSheet("Budget");
      Cell leftBottomCell = sheet.createRow(0).createCell(0);
      Cell rightTopCell = sheet.getRow(0).createCell(1);
      leftBottomCell.setCellValue("LeftBottom");
      rightTopCell.setCellValue("RightTop");

      leftBottomCell.setCellStyle(
          styleRegistry.mergedStyle(
              leftBottomCell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "PATH",
                          null,
                          0.1d,
                          null,
                          null,
                          0.4d,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                              new ExcelGradientStop(1.0d, new ExcelColor(null, null, 10, null))))),
                  null,
                  null)));
      rightTopCell.setCellStyle(
          styleRegistry.mergedStyle(
              rightTopCell,
              new ExcelCellStyle(
                  null,
                  null,
                  null,
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          "PATH",
                          null,
                          null,
                          0.2d,
                          0.3d,
                          null,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#223344")),
                              new ExcelGradientStop(1.0d, new ExcelColor(null, 3, null, null))))),
                  null,
                  null)));

      var leftBottomGradient =
          workbook
              .getStylesSource()
              .getFillAt(
                  (int) ((XSSFCellStyle) leftBottomCell.getCellStyle()).getCoreXf().getFillId())
              .getCTFill()
              .getGradientFill();
      var rightTopGradient =
          workbook
              .getStylesSource()
              .getFillAt(
                  (int) ((XSSFCellStyle) rightTopCell.getCellStyle()).getCoreXf().getFillId())
              .getCTFill()
              .getGradientFill();

      assertTrue(leftBottomGradient.isSetLeft());
      assertFalse(leftBottomGradient.isSetRight());
      assertFalse(leftBottomGradient.isSetTop());
      assertTrue(leftBottomGradient.isSetBottom());
      assertFalse(rightTopGradient.isSetLeft());
      assertTrue(rightTopGradient.isSetRight());
      assertTrue(rightTopGradient.isSetTop());
      assertFalse(rightTopGradient.isSetBottom());
    }
  }

  @Test
  void linearGradientsWithoutDegreeWriteNoGeometryAttributes() throws Exception {
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
                  new ExcelCellFill(
                      null,
                      null,
                      null,
                      new ExcelGradientFill(
                          null,
                          null,
                          null,
                          null,
                          null,
                          null,
                          List.of(
                              new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                              new ExcelGradientStop(1.0d, new ExcelColor("#445566"))))),
                  null,
                  null)));

      var gradient =
          workbook
              .getStylesSource()
              .getFillAt((int) ((XSSFCellStyle) cell.getCellStyle()).getCoreXf().getFillId())
              .getCTFill()
              .getGradientFill();

      assertFalse(gradient.isSetType());
      assertFalse(gradient.isSetDegree());
      assertFalse(gradient.isSetLeft());
      assertFalse(gradient.isSetRight());
      assertFalse(gradient.isSetTop());
      assertFalse(gradient.isSetBottom());
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
    assertEquals(rgb(foregroundColor), snapshot.foregroundColor());
    if (pattern == ExcelFillPattern.SOLID) {
      assertNull(snapshot.backgroundColor());
      return;
    }
    assertEquals(rgb(backgroundColor), snapshot.backgroundColor());
  }

  private static ExcelColorSnapshot rgb(String rgb) {
    return new ExcelColorSnapshot(rgb);
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
