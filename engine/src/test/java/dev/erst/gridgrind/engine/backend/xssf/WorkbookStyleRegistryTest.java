package dev.erst.gridgrind.engine.backend.xssf;

import static dev.erst.gridgrind.engine.backend.xssf.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.junit.jupiter.api.Test;

/** Tests for StyleAdapter utilities, merge behavior, and style snapshots. */
class WorkbookStyleRegistryTest {
  @Test
  void equivalentPointAndTwipPatchesReuseStylesAndFonts() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter registry = new StyleAdapter(workbook);
      var row = workbook.createSheet("Budget").createRow(0);
      var points = new FontHeightSpec.Points(new BigDecimal("13.0"));
      CellFontPatch pointFont =
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(points),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      CellFontPatch twipFont =
          new CellFontPatch(
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(new FontHeightSpec.Twips(260)),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      CellStylePatch pointPatch =
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(pointFont),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      CellStylePatch twipPatch =
          new CellStylePatch(
              Optional.empty(),
              Optional.empty(),
              Optional.of(twipFont),
              Optional.empty(),
              Optional.empty(),
              Optional.empty());
      int initialStyles = workbook.getNumCellStyles();
      int initialFonts = workbook.getNumberOfFonts();

      var pointStyle = registry.mergedStyle(row.createCell(0), pointPatch);
      var twipStyle = registry.mergedStyle(row.createCell(1), twipPatch);

      assertSame(pointStyle, twipStyle);
      assertEquals(initialStyles + 1, workbook.getNumCellStyles());
      assertEquals(initialFonts + 1, workbook.getNumberOfFonts());
      assertSame(points, pointPatch.font().orElseThrow().fontHeight().orElseThrow());
      assertEquals(new BigDecimal("13.0"), points.points());
    }
  }

  @Test
  void resolveNumberFormat_returnsGeneralForNullOrBlank() {
    assertEquals("General", StyleAdapter.resolveNumberFormat(null));
    assertEquals("General", StyleAdapter.resolveNumberFormat(""));
    assertEquals("General", StyleAdapter.resolveNumberFormat("   "));
  }

  @Test
  void resolveNumberFormat_returnsFormatStringWhenPopulated() {
    assertEquals("#,##0.00", StyleAdapter.resolveNumberFormat("#,##0.00"));
    assertEquals("yyyy-mm-dd", StyleAdapter.resolveNumberFormat("yyyy-mm-dd"));
  }

  @Test
  void mergedStyle_preservesUnspecifiedFontAttributesFromExistingCellStyle() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(null, null, font(null, null, null, null, null, true, null), null, null, null)));

      CellStyleReport snapshot = styleRegistry.snapshot(cell);
      assertEquals("Aptos", snapshot.font().fontName());
      assertEquals(300, snapshot.font().fontHeight().twips());
      assertEquals(new BigDecimal("15"), snapshot.font().fontHeight().points());
      assertEquals(rgb("#112233"), snapshot.font().fontColor().orElseThrow());
      assertTrue(snapshot.font().underline());
      assertTrue(snapshot.font().strikeout());
      assertEquals(rgb("#FFF2CC"), fillForegroundColor(snapshot.fill()));
      assertEquals(ExcelBorderStyle.MEDIUM, borderStyle(snapshot.border().top()));
      assertEquals(ExcelBorderStyle.MEDIUM, borderStyle(snapshot.border().right()));
      assertEquals(ExcelBorderStyle.MEDIUM, borderStyle(snapshot.border().bottom()));
      assertEquals(ExcelBorderStyle.MEDIUM, borderStyle(snapshot.border().left()));
    }
  }

  @Test
  void mergedStyle_appliesBorderDefaultsAndExplicitSideOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.patternForeground(
                      ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#DDEBF7"))),
                  new CellBorderPatch(
                      Optional.ofNullable(borderSide(ExcelBorderStyle.THIN)),
                      Optional.empty(),
                      Optional.ofNullable(borderSide(ExcelBorderStyle.DOUBLE)),
                      Optional.empty(),
                      Optional.empty()),
                  null)));

      CellStyleReport snapshot = styleRegistry.snapshot(cell);
      assertEquals(rgb("#DDEBF7"), fillForegroundColor(snapshot.fill()));
      assertEquals(ExcelBorderStyle.THIN, borderStyle(snapshot.border().top()));
      assertEquals(ExcelBorderStyle.DOUBLE, borderStyle(snapshot.border().right()));
      assertEquals(ExcelBorderStyle.THIN, borderStyle(snapshot.border().bottom()));
      assertEquals(ExcelBorderStyle.THIN, borderStyle(snapshot.border().left()));
    }
  }

  @Test
  void mergedStyle_handlesFontPatchVariantsAndPartialBorderOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  font(null, null, "Aptos Display", null, null, false, null),
                  null,
                  new CellBorderPatch(
                      Optional.empty(),
                      Optional.ofNullable(borderSide(ExcelBorderStyle.THICK)),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()),
                  null)));
      CellStyleReport fontNameSnapshot = styleRegistry.snapshot(cell);
      assertEquals("Aptos Display", fontNameSnapshot.font().fontName());
      assertFalse(fontNameSnapshot.font().underline());
      assertEquals(ExcelBorderStyle.THICK, borderStyle(fontNameSnapshot.border().top()));
      assertEquals(ExcelBorderStyle.MEDIUM, borderStyle(fontNameSnapshot.border().right()));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  font(
                      null,
                      null,
                      null,
                      FontHeightReport.fromPoints(new BigDecimal("11.5")),
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
              style(
                  null,
                  null,
                  font(null, null, null, null, ColorSpec.rgb(new RgbHex("#445566")), null, null),
                  null,
                  null,
                  null)));
      assertEquals(rgb("#445566"), styleRegistry.snapshot(cell).font().fontColor().orElseThrow());
    }
  }

  @Test
  void mergedStyle_preservesThemeIndexedAndTintFontColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  font(null, true, null, null, ColorSpec.theme(6, -0.35d), null, null),
                  null,
                  null,
                  null)));
      assertEquals(
          CellColorReport.theme(6, -0.35d),
          styleRegistry.snapshot(cell).font().fontColor().orElseThrow());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  font(
                      null,
                      null,
                      null,
                      null,
                      ColorSpec.indexed(Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex())),
                      null,
                      null),
                  null,
                  null,
                  null)));
      assertEquals(
          CellColorReport.indexed(Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex())),
          styleRegistry.snapshot(cell).font().fontColor().orElseThrow());
    }
  }

  @Test
  void mergedStyle_appliesGradientFillPatchesAndCollapsesExtraFontColorNodes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  font(null, null, null, null, ColorSpec.theme(6, -0.35d), null, null),
                  CellFillSpec.gradient(
                      CellGradientFillSpec.path(
                          Optional.of(0.1d),
                          Optional.of(0.2d),
                          Optional.of(0.3d),
                          Optional.of(0.4d),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                              new CellGradientStopSpec(1.0d, ColorSpec.indexed(5))))),
                  null,
                  null)));

      XSSFCellStyle snapshotStyle = (XSSFCellStyle) cell.getCellStyle();
      XSSFFont snapshotFont = workbook.getFontAt(snapshotStyle.getFontIndex());
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
          assertInstanceOf(
                  CellColorReport.Rgb.class,
                  ColorAdapter.snapshot(workbook, gradientFill.getStopArray(0).getColor())
                      .orElseThrow())
              .rgb()
              .value());
      assertEquals(
          5,
          assertInstanceOf(
                  CellColorReport.Indexed.class,
                  ColorAdapter.snapshot(workbook, gradientFill.getStopArray(1).getColor())
                      .orElseThrow())
              .indexed());
    }
  }

  @Test
  void mergedStyle_collapsesDuplicateFontColorsAndSupportsLinearGradientDefaults()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
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
              style(
                  null,
                  null,
                  font(null, null, null, null, ColorSpec.rgb(new RgbHex("#445566")), null, null),
                  CellFillSpec.gradient(
                      CellGradientFillSpec.linear(
                          Optional.of(37.5d),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                              new CellGradientStopSpec(
                                  1.0d, ColorSpec.rgb(new RgbHex("#445566")))))),
                  null,
                  null)));

      XSSFCellStyle mergedStyle = (XSSFCellStyle) cell.getCellStyle();
      XSSFFont mergedFont = workbook.getFontAt(mergedStyle.getFontIndex());
      var gradientFill =
          workbook
              .getStylesSource()
              .getFillAt((int) mergedStyle.getCoreXf().getFillId())
              .getCTFill()
              .getGradientFill();

      assertEquals(1, mergedFont.getCTFont().sizeOfColorArray());
      assertEquals(
          "#445566",
          assertInstanceOf(
                  CellColorReport.Rgb.class,
                  ColorAdapter.snapshot(workbook, mergedFont.getCTFont().getColorArray(0))
                      .orElseThrow())
              .rgb()
              .value());
      assertNotNull(gradientFill);
      assertFalse(gradientFill.isSetType());
      assertEquals(37.5d, gradientFill.getDegree());
    }
  }

  @Test
  void mergedStyle_appliesFontColorWhenBaseFontHasNoColorNode() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
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
              style(
                  null,
                  null,
                  font(null, null, null, null, ColorSpec.rgb(new RgbHex("#102030")), null, null),
                  null,
                  null,
                  null)));

      XSSFCellStyle mergedStyle = (XSSFCellStyle) cell.getCellStyle();
      XSSFFont mergedFont = workbook.getFontAt(mergedStyle.getFontIndex());
      assertEquals(1, mergedFont.getCTFont().sizeOfColorArray());
      assertEquals(
          "#102030",
          assertInstanceOf(
                  CellColorReport.Rgb.class,
                  ColorAdapter.snapshot(workbook, mergedFont.getCTFont().getColorArray(0))
                      .orElseThrow())
              .rgb()
              .value());
    }
  }

  @Test
  void snapshot_handlesNullAndNonRgbColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      XSSFCellStyle noColorFontStyle = workbook.createCellStyle();
      XSSFFont noColorFont = workbook.createFont();
      noColorFontStyle.setFont(noColorFont);
      cell.setCellStyle(noColorFontStyle);
      assertEquals(Optional.empty(), styleRegistry.snapshot(cell).font().fontColor());

      XSSFCellStyle nonRgbFillStyle = workbook.createCellStyle();
      nonRgbFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      nonRgbFillStyle.setFillForegroundColor(new XSSFColor());
      cell.setCellStyle(nonRgbFillStyle);
      assertThrows(IllegalStateException.class, () -> styleRegistry.snapshot(cell).fill());

      XSSFCellStyle malformedRgbFillStyle = workbook.createCellStyle();
      malformedRgbFillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      malformedRgbFillStyle.setFillForegroundColor(new XSSFColor(new byte[] {0x11, 0x22}));
      cell.setCellStyle(malformedRgbFillStyle);
      assertThrows(IllegalStateException.class, () -> styleRegistry.snapshot(cell).fill());
    }
  }

  @Test
  void invisibleBorderReportPreservesStoredWorkbookColor() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      XSSFCellStyle style = workbook.createCellStyle();
      XSSFColor indexedColor = new XSSFColor();
      indexedColor.setIndexed(IndexedColors.DARK_RED.getIndex());
      style.setBottomBorderColor(indexedColor);
      cell.setCellStyle(style);

      CellStyleReport snapshot = styleRegistry.snapshot(cell);

      assertInstanceOf(CellBorderSideReport.None.class, snapshot.border().bottom());
      assertEquals(
          Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()),
          Short.toUnsignedInt(style.getBottomBorderXSSFColor().getIndexed()));
    }
  }

  @Test
  void mergedStyle_tracksExpandedB4StyleDepth() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  alignment(
                      true, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP, 45, 3),
                  font(
                      true,
                      false,
                      "Aptos",
                      FontHeightReport.fromPoints(new BigDecimal("11.5")),
                      ColorSpec.rgb(new RgbHex("#102030")),
                      true,
                      false),
                  CellFillSpec.patternColors(
                      ExcelFillPattern.THIN_HORIZONTAL_BANDS,
                      ColorSpec.rgb(new RgbHex("#FFF2CC")),
                      ColorSpec.rgb(new RgbHex("#DDEBF7"))),
                  new CellBorderPatch(
                      Optional.ofNullable(
                          borderSide(ExcelBorderStyle.THIN, ColorSpec.rgb(new RgbHex("#203040")))),
                      Optional.empty(),
                      Optional.ofNullable(
                          borderSide(
                              ExcelBorderStyle.DOUBLE, ColorSpec.rgb(new RgbHex("#304050")))),
                      Optional.empty(),
                      Optional.empty()),
                  new CellProtectionPatch(Optional.of(false), Optional.of(true)))));

      CellStyleReport snapshot = styleRegistry.snapshot(cell);
      assertTrue(snapshot.alignment().wrapText());
      assertEquals(45, snapshot.alignment().textRotation());
      assertEquals(3, snapshot.alignment().indentation());
      assertEquals(rgb("#102030"), snapshot.font().fontColor().orElseThrow());
      assertEquals(ExcelFillPattern.THIN_HORIZONTAL_BANDS, fillPattern(snapshot.fill()));
      assertEquals(rgb("#FFF2CC"), fillForegroundColor(snapshot.fill()));
      assertEquals(rgb("#DDEBF7"), fillBackgroundColor(snapshot.fill()));
      assertEquals(rgb("#203040"), borderColor(snapshot.border().top()));
      assertEquals(rgb("#304050"), borderColor(snapshot.border().right()));
      assertFalse(snapshot.protection().locked());
      assertTrue(snapshot.protection().hiddenFormula());
    }
  }

  @Test
  void defaultStylesAndDateFactoriesPreserveNonFormatState() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(styledBaseCellStyle(workbook));

      assertEquals(workbook.getCellStyleAt(0).getIndex(), styleRegistry.defaultStyle().getIndex());
      assertEquals("General", styleRegistry.defaultSnapshot().numberFormat());

      cell.setCellStyle(styleRegistry.localDateStyle(cell));
      CellStyleReport dateSnapshot = styleRegistry.snapshot(cell);
      assertEquals("yyyy-mm-dd", dateSnapshot.numberFormat());
      assertEquals("Aptos", dateSnapshot.font().fontName());
      assertEquals(rgb("#FFF2CC"), fillForegroundColor(dateSnapshot.fill()));

      cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      CellStyleReport dateTimeSnapshot = styleRegistry.snapshot(cell);
      assertEquals("yyyy-mm-dd hh:mm:ss", dateTimeSnapshot.numberFormat());
      assertEquals(ExcelBorderStyle.MEDIUM, borderStyle(dateTimeSnapshot.border().top()));
      assertEquals(rgb("#112233"), dateTimeSnapshot.font().fontColor().orElseThrow());
    }
  }

  @Test
  void mergedStyle_handlesForegroundOnlyFillPatchesAndFillClears() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternForeground(
                      ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#ABC123"))))));
      assertFillSnapshot(
          styleRegistry.snapshot(cell).fill(), ExcelFillPattern.SOLID, "#ABC123", null);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternColors(
                      ExcelFillPattern.BRICKS,
                      ColorSpec.rgb(new RgbHex("#112233")),
                      ColorSpec.rgb(new RgbHex("#445566"))))));
      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternForeground(
                      ExcelFillPattern.BRICKS, ColorSpec.rgb(new RgbHex("#AA5500"))))));
      assertFillSnapshot(
          styleRegistry.snapshot(cell).fill(), ExcelFillPattern.BRICKS, "#AA5500", "#445566");

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternBackground(
                      ExcelFillPattern.BRICKS, ColorSpec.rgb(new RgbHex("#556677"))))));
      assertFillSnapshot(
          styleRegistry.snapshot(cell).fill(), ExcelFillPattern.BRICKS, "#AA5500", "#556677");

      cell.setCellStyle(
          styleRegistry.mergedStyle(cell, fillPatch(CellFillSpec.pattern(ExcelFillPattern.NONE))));
      assertFillSnapshot(styleRegistry.snapshot(cell).fill(), ExcelFillPattern.NONE, null, null);
    }
  }

  @Test
  void mergedStyle_mergesBorderDefaultsColorsAndExplicitOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new CellBorderPatch(
                      Optional.ofNullable(
                          borderSide(ExcelBorderStyle.THIN, ColorSpec.rgb(new RgbHex("#112233")))),
                      Optional.ofNullable(borderSide(null, ColorSpec.rgb(new RgbHex("#445566")))),
                      Optional.ofNullable(borderSide(ExcelBorderStyle.DOUBLE, null)),
                      Optional.empty(),
                      Optional.empty()))));
      CellStyleReport mergedBorderSnapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.THIN, borderStyle(mergedBorderSnapshot.border().top()));
      assertEquals(rgb("#445566"), borderColor(mergedBorderSnapshot.border().top()));
      assertEquals(ExcelBorderStyle.DOUBLE, borderStyle(mergedBorderSnapshot.border().right()));
      assertEquals(rgb("#112233"), borderColor(mergedBorderSnapshot.border().right()));
      assertEquals(rgb("#112233"), borderColor(mergedBorderSnapshot.border().bottom()));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new CellBorderPatch(
                      Optional.empty(),
                      Optional.ofNullable(borderSide(ExcelBorderStyle.NONE)),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()))));
      CellStyleReport clearedBorderSnapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.NONE, borderStyle(clearedBorderSnapshot.border().top()));
      assertNull(borderColor(clearedBorderSnapshot.border().top()));
      assertEquals(ExcelBorderStyle.DOUBLE, borderStyle(clearedBorderSnapshot.border().right()));
      assertEquals(rgb("#112233"), borderColor(clearedBorderSnapshot.border().right()));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  cell,
                  borderPatch(
                      new CellBorderPatch(
                          Optional.empty(),
                          Optional.ofNullable(
                              borderSide(null, ColorSpec.rgb(new RgbHex("#778899")))),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))));

      Cell blankCell = workbook.getSheet("Budget").createRow(1).createCell(0);
      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  blankCell,
                  borderPatch(
                      new CellBorderPatch(
                          Optional.empty(),
                          Optional.ofNullable(
                              borderSide(null, ColorSpec.rgb(new RgbHex("#99AABB")))),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              styleRegistry.mergedStyle(
                  blankCell,
                  borderPatch(
                      new CellBorderPatch(
                          Optional.ofNullable(borderSide(ExcelBorderStyle.NONE)),
                          Optional.ofNullable(
                              borderSide(null, ColorSpec.rgb(new RgbHex("#CC8844")))),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))));
    }
  }

  @Test
  void snapshotReportsInvisibleBorderWithoutItsStoredColor() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFCellStyle style = workbook.createCellStyle();
      style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.NONE);
      style.setTopBorderColor(new XSSFColor(new byte[] {17, 34, 51}, null));
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellStyle(style);

      assertNotNull(style.getTopBorderXSSFColor());
      assertInstanceOf(
          CellBorderSideReport.None.class,
          new StyleAdapter(workbook).snapshot(cell).border().top());

      style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
      CellBorderSideReport.Colored visible =
          assertInstanceOf(
              CellBorderSideReport.Colored.class,
              new StyleAdapter(workbook).snapshot(cell).border().top());
      assertEquals(rgb("#112233"), visible.color());
    }
  }

  @Test
  void mergedStyle_clearsNoneBordersWithoutCrashingWhenNoColorExists() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new CellBorderPatch(
                      Optional.ofNullable(borderSide(ExcelBorderStyle.NONE)),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()))));

      CellStyleReport snapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.NONE, borderStyle(snapshot.border().top()));
      assertNull(borderColor(snapshot.border().top()));
      assertEquals(ExcelBorderStyle.NONE, borderStyle(snapshot.border().right()));
      assertNull(borderColor(snapshot.border().right()));
      assertEquals(ExcelBorderStyle.NONE, borderStyle(snapshot.border().bottom()));
      assertNull(borderColor(snapshot.border().bottom()));
      assertEquals(ExcelBorderStyle.NONE, borderStyle(snapshot.border().left()));
      assertNull(borderColor(snapshot.border().left()));
    }
  }

  @Test
  void mergedStyle_clearsExistingBorderColorsWhenAStyleIsResetToNone() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new CellBorderPatch(
                      Optional.ofNullable(
                          borderSide(ExcelBorderStyle.THIN, ColorSpec.rgb(new RgbHex("#112233")))),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()))));
      assertEquals(rgb("#112233"), borderColor(styleRegistry.snapshot(cell).border().top()));

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              borderPatch(
                  new CellBorderPatch(
                      Optional.ofNullable(borderSide(ExcelBorderStyle.NONE)),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()))));

      CellStyleReport snapshot = styleRegistry.snapshot(cell);
      assertEquals(ExcelBorderStyle.NONE, borderStyle(snapshot.border().top()));
      assertNull(borderColor(snapshot.border().top()));
    }
  }

  @Test
  void mergedStyle_appliesProtectionBranchesIndependently() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              protectionPatch(new CellProtectionPatch(Optional.of(false), Optional.empty()))));
      CellStyleReport lockedSnapshot = styleRegistry.snapshot(cell);
      assertFalse(lockedSnapshot.protection().locked());
      assertFalse(lockedSnapshot.protection().hiddenFormula());

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell, protectionPatch(new CellProtectionPatch(Optional.empty(), Optional.of(true)))));
      CellStyleReport hiddenSnapshot = styleRegistry.snapshot(cell);
      assertFalse(hiddenSnapshot.protection().locked());
      assertTrue(hiddenSnapshot.protection().hiddenFormula());
    }
  }

  @Test
  void snapshot_roundTripsEverySupportedFillPattern() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      var sheet = workbook.createSheet("Patterns");

      ExcelFillPattern[] patterns = ExcelFillPattern.values();
      for (int columnIndex = 0; columnIndex < patterns.length; columnIndex++) {
        ExcelFillPattern pattern = patterns[columnIndex];
        Cell cell = sheet.createRow(0).createCell(columnIndex);
        cell.setCellStyle(styleRegistry.mergedStyle(cell, fillPatch(fillFor(pattern))));
        CellFillReport snapshot = styleRegistry.snapshot(cell).fill();
        assertFillSnapshot(snapshot, pattern, "#102030", "#405060");
      }
    }
  }

  @Test
  void snapshot_readsBackgroundOnlyPatternFillVariants() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternBackground(ExcelFillPattern.BRICKS, ColorSpec.indexed(10)))));

      CellFillReport snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.BRICKS, fillPattern(snapshot));
      assertNull(fillForegroundColor(snapshot));
      assertEquals(CellColorReport.indexed(10), fillBackgroundColor(snapshot));
    }
  }

  @Test
  void snapshot_readsPatternWithoutColorsAsPatternOnly() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell, fillPatch(CellFillSpec.pattern(ExcelFillPattern.BRICKS))));

      CellFillReport snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.BRICKS, fillPattern(snapshot));
      assertNull(fillForegroundColor(snapshot));
      assertNull(fillBackgroundColor(snapshot));
    }
  }

  @Test
  void mergedStyle_clearsBackgroundWhenSolidForegroundOverridesPatternedFill() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternColors(
                      ExcelFillPattern.BRICKS,
                      ColorSpec.rgb(new RgbHex("#102030")),
                      ColorSpec.rgb(new RgbHex("#405060"))))));
      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternForeground(
                      ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#AABBCC"))))));

      CellFillReport snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.SOLID, fillPattern(snapshot));
      assertEquals(CellColorReport.rgb(new RgbHex("#AABBCC")), fillForegroundColor(snapshot));
      assertNull(fillBackgroundColor(snapshot));
    }
  }

  @Test
  void mergedStyle_patternOnlySolidClearsInheritedBackgroundColor() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternColors(
                      ExcelFillPattern.BRICKS,
                      ColorSpec.rgb(new RgbHex("#102030")),
                      ColorSpec.rgb(new RgbHex("#405060"))))));
      cell.setCellStyle(
          styleRegistry.mergedStyle(cell, fillPatch(CellFillSpec.pattern(ExcelFillPattern.SOLID))));

      CellFillReport snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.SOLID, fillPattern(snapshot));
      assertEquals(CellColorReport.rgb(new RgbHex("#102030")), fillForegroundColor(snapshot));
      assertNull(fillBackgroundColor(snapshot));
    }
  }

  @Test
  void mergedStyle_patternOnlyNonSolidPreservesInheritedFillColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              fillPatch(
                  CellFillSpec.patternColors(
                      ExcelFillPattern.BRICKS,
                      ColorSpec.rgb(new RgbHex("#102030")),
                      ColorSpec.rgb(new RgbHex("#405060"))))));
      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell, fillPatch(CellFillSpec.pattern(ExcelFillPattern.FINE_DOTS))));

      CellFillReport snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.FINE_DOTS, fillPattern(snapshot));
      assertEquals(CellColorReport.rgb(new RgbHex("#102030")), fillForegroundColor(snapshot));
      assertEquals(CellColorReport.rgb(new RgbHex("#405060")), fillBackgroundColor(snapshot));
    }
  }

  @Test
  void snapshot_readsGradientFillStopsAndWorkbookColorSemantics() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
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

      CellFillReport snapshot = styleRegistry.snapshot(cell).fill();

      assertEquals(ExcelFillPattern.NONE, fillPattern(snapshot));
      CellGradientFillReport gradient = fillGradient(snapshot);
      assertNotNull(gradient);
      assertEquals("LINEAR", gradientType(gradient));
      assertEquals(45.0d, gradientDegree(gradient));
      assertEquals(2, gradient.stops().size());
      assertEquals(CellColorReport.rgb(new RgbHex("#112233")), gradient.stops().get(0).color());
      assertEquals(CellColorReport.indexed(10), gradient.stops().get(1).color());
    }
  }

  @Test
  void gradientFillSnapshotReflectsExplicitPathOffsets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
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

      CellGradientFillReport snapshot = styleRegistry.gradientFillSnapshot(gradient);

      assertEquals("PATH", gradientType(snapshot));
      assertNull(gradientDegree(snapshot));
      assertEquals(0.1d, gradientLeft(snapshot));
      assertEquals(0.2d, gradientRight(snapshot));
      assertEquals(0.3d, gradientTop(snapshot));
      assertEquals(0.4d, gradientBottom(snapshot));
      assertEquals(CellColorReport.theme(2), snapshot.stops().get(0).color());
      assertEquals(CellColorReport.indexed(5), snapshot.stops().get(1).color());
    }
  }

  @Test
  void gradientFillSnapshotInfersPathTypeWhenOffsetsExistWithoutExplicitType() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
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

      CellGradientFillReport snapshot = styleRegistry.gradientFillSnapshot(gradient);

      assertEquals("PATH", gradientType(snapshot));
      assertNull(gradientDegree(snapshot));
      assertEquals(0.1d, gradientLeft(snapshot));
      assertEquals(0.2d, gradientRight(snapshot));
      assertEquals(0.3d, gradientTop(snapshot));
      assertEquals(0.4d, gradientBottom(snapshot));
    }
  }

  @Test
  void gradientFillSnapshotDefaultsLinearDegreeToNullWhenAbsent() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      var gradient =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill.Factory.newInstance();
      var firstStop = gradient.addNewStop();
      firstStop.setPosition(0.0d);
      firstStop.addNewColor().setRgb(new byte[] {(byte) 0xFF, 0x11, 0x22, 0x33});
      var secondStop = gradient.addNewStop();
      secondStop.setPosition(1.0d);
      secondStop.addNewColor().setIndexed(10L);

      CellGradientFillReport snapshot = styleRegistry.gradientFillSnapshot(gradient);

      assertEquals("LINEAR", gradientType(snapshot));
      assertNull(gradientDegree(snapshot));
      assertEquals(CellColorReport.rgb(new RgbHex("#112233")), snapshot.stops().get(0).color());
      assertEquals(CellColorReport.indexed(10), snapshot.stops().get(1).color());
    }
  }

  @Test
  void gradientFillSnapshotRejectsStopsWithoutColorDefinition() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      var gradient =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill.Factory.newInstance();
      var stop = gradient.addNewStop();
      stop.setPosition(0.5d);

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class, () -> styleRegistry.gradientFillSnapshot(gradient));

      assertEquals(
          "Gradient stop at position 0.5 is missing its color definition.", failure.getMessage());
    }
  }

  @Test
  void snapshotReportsPathGradientGeometryForStyledCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);
      cell.setCellValue("Path");

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(
                      CellGradientFillSpec.path(
                          Optional.of(0.1d),
                          Optional.of(0.2d),
                          Optional.of(0.3d),
                          Optional.of(0.4d),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                              new CellGradientStopSpec(1.0d, ColorSpec.indexed(10))))),
                  null,
                  new CellProtectionPatch(Optional.of(false), Optional.of(true)))));

      CellStyleReport snapshot = styleRegistry.snapshot(cell);

      CellGradientFillReport gradient = fillGradient(snapshot.fill());
      assertEquals("PATH", gradientType(gradient));
      assertNull(gradientDegree(gradient));
      assertEquals(0.1d, gradientLeft(gradient));
      assertEquals(0.2d, gradientRight(gradient));
      assertEquals(0.3d, gradientTop(gradient));
      assertEquals(0.4d, gradientBottom(gradient));
      assertFalse(snapshot.protection().locked());
      assertTrue(snapshot.protection().hiddenFormula());
    }
  }

  @Test
  void distinctGradientFillsDoNotAliasWhenWorkbookUsesMultipleGradientStyles() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      var sheet = workbook.createSheet("Budget");
      Cell linearCell = sheet.createRow(0).createCell(0);
      Cell pathCell = sheet.getRow(0).createCell(1);
      linearCell.setCellValue("Linear");
      pathCell.setCellValue("Path");

      linearCell.setCellStyle(
          styleRegistry.mergedStyle(
              linearCell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(
                      CellGradientFillSpec.linear(
                          Optional.of(45.0d),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#1F497D"))),
                              new CellGradientStopSpec(1.0d, ColorSpec.theme(4, 0.45d))))),
                  null,
                  new CellProtectionPatch(Optional.of(true), Optional.of(true)))));
      pathCell.setCellStyle(
          styleRegistry.mergedStyle(
              pathCell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(
                      CellGradientFillSpec.path(
                          Optional.of(0.1d),
                          Optional.of(0.2d),
                          Optional.of(0.3d),
                          Optional.of(0.4d),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                              new CellGradientStopSpec(1.0d, ColorSpec.indexed(10))))),
                  null,
                  new CellProtectionPatch(Optional.of(false), Optional.of(true)))));

      CellStyleReport linearSnapshot = styleRegistry.snapshot(linearCell);
      CellStyleReport pathSnapshot = styleRegistry.snapshot(pathCell);

      assertEquals("LINEAR", gradientType(fillGradient(linearSnapshot.fill())));
      assertEquals(45.0d, gradientDegree(fillGradient(linearSnapshot.fill())));
      assertEquals("PATH", gradientType(fillGradient(pathSnapshot.fill())));
      assertEquals(0.1d, gradientLeft(fillGradient(pathSnapshot.fill())));
      assertEquals(0.2d, gradientRight(fillGradient(pathSnapshot.fill())));
      assertEquals(0.3d, gradientTop(fillGradient(pathSnapshot.fill())));
      assertEquals(0.4d, gradientBottom(fillGradient(pathSnapshot.fill())));
      assertFalse(pathSnapshot.protection().locked());
      assertTrue(pathSnapshot.protection().hiddenFormula());
    }
  }

  @Test
  void identicalGradientFillsReuseOneStylesTableFillId() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      var sheet = workbook.createSheet("Budget");
      Cell firstCell = sheet.createRow(0).createCell(0);
      Cell secondCell = sheet.getRow(0).createCell(1);
      firstCell.setCellValue("Locked");
      secondCell.setCellValue("Unlocked");
      CellGradientFillSpec sharedGradient =
          CellGradientFillSpec.linear(
              Optional.of(45.0d),
              List.of(
                  new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#1F497D"))),
                  new CellGradientStopSpec(1.0d, ColorSpec.theme(4, 0.45d))));

      firstCell.setCellStyle(
          styleRegistry.mergedStyle(
              firstCell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(sharedGradient),
                  null,
                  new CellProtectionPatch(Optional.of(true), Optional.of(true)))));
      secondCell.setCellStyle(
          styleRegistry.mergedStyle(
              secondCell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(sharedGradient),
                  null,
                  new CellProtectionPatch(Optional.of(false), Optional.of(true)))));

      long firstFillId = ((XSSFCellStyle) firstCell.getCellStyle()).getCoreXf().getFillId();
      long secondFillId = ((XSSFCellStyle) secondCell.getCellStyle()).getCoreXf().getFillId();

      assertEquals(firstFillId, secondFillId);
    }
  }

  @Test
  void pathGradientsWriteOnlyProvidedOffsets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      var sheet = workbook.createSheet("Budget");
      Cell leftBottomCell = sheet.createRow(0).createCell(0);
      Cell rightTopCell = sheet.getRow(0).createCell(1);
      leftBottomCell.setCellValue("LeftBottom");
      rightTopCell.setCellValue("RightTop");

      leftBottomCell.setCellStyle(
          styleRegistry.mergedStyle(
              leftBottomCell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(
                      CellGradientFillSpec.path(
                          Optional.of(0.1d),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.of(0.4d),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                              new CellGradientStopSpec(1.0d, ColorSpec.indexed(10))))),
                  null,
                  null)));
      rightTopCell.setCellStyle(
          styleRegistry.mergedStyle(
              rightTopCell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(
                      CellGradientFillSpec.path(
                          Optional.empty(),
                          Optional.of(0.2d),
                          Optional.of(0.3d),
                          Optional.empty(),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#223344"))),
                              new CellGradientStopSpec(1.0d, ColorSpec.theme(3))))),
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
      StyleAdapter styleRegistry = new StyleAdapter(workbook);
      Cell cell = workbook.createSheet("Budget").createRow(0).createCell(0);

      cell.setCellStyle(
          styleRegistry.mergedStyle(
              cell,
              style(
                  null,
                  null,
                  null,
                  CellFillSpec.gradient(
                      CellGradientFillSpec.linear(
                          Optional.empty(),
                          List.of(
                              new CellGradientStopSpec(0.0d, ColorSpec.rgb(new RgbHex("#112233"))),
                              new CellGradientStopSpec(
                                  1.0d, ColorSpec.rgb(new RgbHex("#445566")))))),
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

  private CellStylePatch style(
      String numberFormat,
      CellAlignmentPatch alignment,
      CellFontPatch font,
      CellFillSpec fill,
      CellBorderPatch border,
      CellProtectionPatch protection) {
    return new CellStylePatch(
        Optional.ofNullable(numberFormat),
        Optional.ofNullable(alignment),
        Optional.ofNullable(font),
        Optional.ofNullable(fill),
        Optional.ofNullable(border),
        Optional.ofNullable(protection));
  }

  private CellFontPatch font(
      Boolean bold,
      Boolean italic,
      String fontName,
      FontHeightReport fontHeight,
      ColorSpec fontColor,
      Boolean underline,
      Boolean strikeout) {
    return new CellFontPatch(
        Optional.ofNullable(bold),
        Optional.ofNullable(italic),
        Optional.ofNullable(fontName),
        Optional.ofNullable(fontHeight).map(height -> new FontHeightSpec.Twips(height.twips())),
        Optional.ofNullable(fontColor),
        Optional.ofNullable(underline),
        Optional.ofNullable(strikeout));
  }

  private CellAlignmentPatch alignment(
      Boolean wrapText,
      ExcelHorizontalAlignment horizontalAlignment,
      ExcelVerticalAlignment verticalAlignment,
      Integer textRotation,
      Integer indentation) {
    return new CellAlignmentPatch(
        Optional.ofNullable(wrapText),
        Optional.ofNullable(horizontalAlignment),
        Optional.ofNullable(verticalAlignment),
        Optional.ofNullable(textRotation),
        Optional.ofNullable(indentation));
  }

  private BorderSidePatch borderSide(ExcelBorderStyle style, ColorSpec color) {
    return new BorderSidePatch(Optional.ofNullable(style), Optional.ofNullable(color));
  }

  private BorderSidePatch borderSide(ExcelBorderStyle style) {
    return borderSide(style, null);
  }

  private CellStylePatch fillPatch(CellFillSpec fill) {
    return style(null, null, null, fill, null, null);
  }

  private CellStylePatch borderPatch(CellBorderPatch border) {
    return style(null, null, null, null, border, null);
  }

  private CellStylePatch protectionPatch(CellProtectionPatch protection) {
    return style(null, null, null, null, null, protection);
  }

  private CellFillSpec fillFor(ExcelFillPattern pattern) {
    if (pattern == ExcelFillPattern.NONE) {
      return CellFillSpec.pattern(pattern);
    }
    if (pattern == ExcelFillPattern.SOLID) {
      return CellFillSpec.patternForeground(pattern, ColorSpec.rgb(new RgbHex("#102030")));
    }
    return CellFillSpec.patternColors(
        pattern, ColorSpec.rgb(new RgbHex("#102030")), ColorSpec.rgb(new RgbHex("#405060")));
  }

  private void assertFillSnapshot(
      CellFillReport snapshot,
      ExcelFillPattern pattern,
      String foregroundColor,
      String backgroundColor) {
    assertEquals(pattern, fillPattern(snapshot));
    if (pattern == ExcelFillPattern.NONE) {
      assertNull(fillForegroundColor(snapshot));
      assertNull(fillBackgroundColor(snapshot));
      return;
    }
    assertEquals(rgb(foregroundColor), fillForegroundColor(snapshot));
    if (pattern == ExcelFillPattern.SOLID) {
      assertNull(fillBackgroundColor(snapshot));
      return;
    }
    assertEquals(rgb(backgroundColor), fillBackgroundColor(snapshot));
  }

  private static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(new RgbHex(rgb));
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
