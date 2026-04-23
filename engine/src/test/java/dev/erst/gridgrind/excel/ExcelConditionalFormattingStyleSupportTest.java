package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.math.BigDecimal;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;

/** Tests for conditional-formatting differential-style XML read and write support. */
class ExcelConditionalFormattingStyleSupportTest {
  @Test
  void applyStyleAndSnapshotRoundTripPreserveExplicitAttributes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFConditionalFormattingRule rule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("A1>10");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A3")
              },
              rule);
      CTCfRule ctRule = sheet.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0);

      ExcelDifferentialStyle authoredStyle =
          new ExcelDifferentialStyle(
              "0.00",
              true,
              false,
              ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
              "#102030",
              true,
              true,
              "#E0F0AA",
              new ExcelDifferentialBorder(
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060"),
                  null,
                  null,
                  null,
                  null));

      ExcelConditionalFormattingStyleSupport.applyStyle(workbook, ctRule, authoredStyle);

      ExcelDifferentialStyleSnapshot snapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(workbook.getStylesSource(), ctRule);

      assertEquals(
          new ExcelDifferentialStyleSnapshot(
              "0.00",
              true,
              false,
              ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
              "#102030",
              true,
              true,
              "#E0F0AA",
              new ExcelDifferentialBorder(
                  null,
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060"),
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060"),
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060"),
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060")),
              List.of()),
          snapshot);
    }
  }

  @Test
  void snapshotStyleFlagsUnsupportedDifferentialFeatures() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFConditionalFormattingRule rule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("A1>0");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A3")
              },
              rule);
      CTCfRule ctRule = sheet.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0);

      var dxf = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf.Factory.newInstance();
      dxf.addNewAlignment()
          .setHorizontal(
              org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment.CENTER);
      dxf.addNewProtection().setLocked(true);
      var font = dxf.addNewFont();
      font.addNewName().setVal("Aptos");
      font.addNewColor().setTheme(1);
      var fill = dxf.addNewFill().addNewPatternFill();
      fill.setPatternType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.DARK_GRAY);
      fill.addNewBgColor().setIndexed(4);
      var diagonal = dxf.addNewBorder().addNewDiagonal();
      diagonal.setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
      ExcelConditionalFormattingStyleSupport.attachStyle(workbook.getStylesSource(), ctRule, dxf);

      ExcelDifferentialStyleSnapshot snapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(workbook.getStylesSource(), ctRule);

      assertEquals(
          List.of(
              ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES,
              ExcelConditionalFormattingUnsupportedFeature.FILL_BACKGROUND_COLOR,
              ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN,
              ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY,
              ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT,
              ExcelConditionalFormattingUnsupportedFeature.PROTECTION),
          snapshot.unsupportedFeatures());
      assertNull(snapshot.fontColor());
      assertNull(snapshot.fillColor());
      assertNull(snapshot.border());
    }
  }

  @Test
  void snapshotStyleFlagsInvalidDifferentialStyleReference() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFConditionalFormattingRule rule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("A1>0");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A3")
              },
              rule);
      CTCfRule ctRule = sheet.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0);
      ctRule.setDxfId(99L);

      ExcelDifferentialStyleSnapshot snapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(workbook.getStylesSource(), ctRule);

      assertEquals(
          new ExcelDifferentialStyleSnapshot(
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              List.of(ExcelConditionalFormattingUnsupportedFeature.STYLE_REFERENCE)),
          snapshot);
    }
  }

  @Test
  void snapshotStyleFlagsNegativeDifferentialStyleReference() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFConditionalFormattingRule rule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("A1>0");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A3")
              },
              rule);
      CTCfRule ctRule = sheet.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0);
      ctRule.setDxfId(-1L);

      assertEquals(
          new ExcelDifferentialStyleSnapshot(
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              List.of(ExcelConditionalFormattingUnsupportedFeature.STYLE_REFERENCE)),
          ExcelConditionalFormattingStyleSupport.snapshotStyle(workbook.getStylesSource(), ctRule));
    }
  }

  @Test
  void rgbAndPatternHelpersNormalizeRawColorPayloads() {
    var argbColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    argbColor.setRgb(new byte[] {(byte) 0xFF, 0x10, 0x20, 0x30});
    var rgbColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    rgbColor.setRgb(new byte[] {0x40, 0x50, 0x60});
    var unsupportedColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    unsupportedColor.setRgb(new byte[] {0x01, 0x02});

    assertNull(ExcelConditionalFormattingStyleSupport.rgbHexFromCtColor(null));
    assertEquals("#102030", ExcelConditionalFormattingStyleSupport.rgbHexFromCtColor(argbColor));
    assertEquals("#405060", ExcelConditionalFormattingStyleSupport.rgbHexFromCtColor(rgbColor));
    assertNull(ExcelConditionalFormattingStyleSupport.rgbHexFromCtColor(unsupportedColor));
  }

  @Test
  void patternHelpersDifferentiateUnsetSolidAndUnsupportedPatterns() {
    var unsetPattern =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    var nonePattern =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    nonePattern.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.NONE);
    var solidPattern =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    solidPattern.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
    var darkPattern =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    darkPattern.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.DARK_GRID);

    assertFalse(ExcelConditionalFormattingStyleSupport.patternTypeIsUnsupported(unsetPattern));
    assertFalse(ExcelConditionalFormattingStyleSupport.patternTypeIsUnsupported(nonePattern));
    assertFalse(ExcelConditionalFormattingStyleSupport.patternTypeIsUnsupported(solidPattern));
    assertTrue(ExcelConditionalFormattingStyleSupport.patternTypeIsUnsupported(darkPattern));
  }

  @Test
  void borderHelpersCoverMappingsSparseSidesAndBooleanFlags() {
    var emptySide =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
    var coloredSide =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
    coloredSide.setStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.DASH_DOT);
    coloredSide.addNewColor().setRgb(new byte[] {(byte) 0xFF, 0x01, 0x23, 0x45});
    var unsupportedSide =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
    unsupportedSide.setStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
    unsupportedSide.addNewColor().setTheme(1L);

    assertNull(ExcelConditionalFormattingStyleSupport.snapshotBorderSide(null));
    assertNull(ExcelConditionalFormattingStyleSupport.snapshotBorderSide(emptySide));
    assertEquals(
        new ExcelDifferentialBorderSide(ExcelBorderStyle.DASH_DOT, "#012345"),
        ExcelConditionalFormattingStyleSupport.snapshotBorderSide(coloredSide));
    assertNull(ExcelConditionalFormattingStyleSupport.snapshotBorderSide(unsupportedSide));

    var colorOnlySide =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
    colorOnlySide.addNewColor().setRgb(new byte[] {(byte) 0xFF, 0x22, 0x33, 0x44});
    assertEquals(
        new ExcelDifferentialBorderSide(ExcelBorderStyle.NONE, "#223344"),
        ExcelConditionalFormattingStyleSupport.snapshotBorderSide(colorOnlySide));

    for (ExcelBorderStyle borderStyle : ExcelBorderStyle.values()) {
      assertEquals(
          borderStyle,
          ExcelConditionalFormattingStyleSupport.fromCtBorderStyle(
              ExcelConditionalFormattingStyleSupport.toCtBorderStyle(borderStyle).intValue()));
    }
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelConditionalFormattingStyleSupport.fromCtBorderStyle(Integer.MAX_VALUE));

    ExcelDifferentialBorderSide thinSide =
        new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, null);

    var complexBorder =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder.Factory.newInstance();
    assertFalse(ExcelConditionalFormattingStyleSupport.hasComplexBorderFeatures(complexBorder));
    complexBorder.setDiagonalDown(true);
    assertTrue(ExcelConditionalFormattingStyleSupport.hasComplexBorderFeatures(complexBorder));

    var unsupportedReferenceBorder =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder.Factory.newInstance();
    unsupportedReferenceBorder
        .addNewTop()
        .setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
    assertTrue(
        ExcelConditionalFormattingStyleSupport.hasUnsupportedSideReference(
            unsupportedReferenceBorder, null, null, null, null));
    var unsupportedRightBorder =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder.Factory.newInstance();
    unsupportedRightBorder
        .addNewRight()
        .setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
    assertTrue(
        ExcelConditionalFormattingStyleSupport.hasUnsupportedSideReference(
            unsupportedRightBorder, thinSide, null, thinSide, thinSide));
    var unsupportedBottomBorder =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder.Factory.newInstance();
    unsupportedBottomBorder
        .addNewBottom()
        .setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
    assertTrue(
        ExcelConditionalFormattingStyleSupport.hasUnsupportedSideReference(
            unsupportedBottomBorder, thinSide, thinSide, null, thinSide));
    var unsupportedLeftBorder =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder.Factory.newInstance();
    unsupportedLeftBorder
        .addNewLeft()
        .setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
    assertTrue(
        ExcelConditionalFormattingStyleSupport.hasUnsupportedSideReference(
            unsupportedLeftBorder, thinSide, thinSide, thinSide, null));

    assertNull(ExcelConditionalFormattingStyleSupport.borderValue(null, null, null, null));
    assertEquals(
        new ExcelDifferentialBorder(null, thinSide, null, null, null),
        ExcelConditionalFormattingStyleSupport.borderValue(thinSide, null, null, null));
    assertEquals(
        new ExcelDifferentialBorder(null, null, thinSide, null, null),
        ExcelConditionalFormattingStyleSupport.borderValue(null, thinSide, null, null));
    assertEquals(
        new ExcelDifferentialBorder(null, null, null, thinSide, null),
        ExcelConditionalFormattingStyleSupport.borderValue(null, null, thinSide, null));
    assertEquals(
        new ExcelDifferentialBorder(null, null, null, null, thinSide),
        ExcelConditionalFormattingStyleSupport.borderValue(null, null, null, thinSide));

    assertTrue(ExcelConditionalFormattingStyleSupport.underline(null));
    var implicitUnderline =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty.Factory
            .newInstance();
    assertTrue(ExcelConditionalFormattingStyleSupport.underline(implicitUnderline));
    var noUnderline =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty.Factory
            .newInstance();
    noUnderline.setVal(org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues.NONE);
    assertFalse(ExcelConditionalFormattingStyleSupport.underline(noUnderline));
  }

  @Test
  void fontMetadataAndPatternForegroundHelpersReportUnsupportedStates() {
    var font = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont.Factory.newInstance();
    assertFalse(ExcelConditionalFormattingStyleSupport.hasUnsupportedFontAttributes(font));
    font.addNewName().setVal("Aptos");
    assertTrue(ExcelConditionalFormattingStyleSupport.hasUnsupportedFontAttributes(font));

    var dxf = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf.Factory.newInstance();
    assertEquals(
        List.of(), ExcelConditionalFormattingStyleSupport.metadataUnsupportedFeatures(dxf));
    dxf.addNewAlignment()
        .setHorizontal(
            org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment.CENTER);
    dxf.addNewProtection().setLocked(true);
    assertEquals(
        List.of(
            ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT,
            ExcelConditionalFormattingUnsupportedFeature.PROTECTION),
        ExcelConditionalFormattingStyleSupport.metadataUnsupportedFeatures(dxf));

    assertNull(ExcelConditionalFormattingStyleSupport.booleanProperty(null));
    var implicitTrue =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty.Factory.newInstance();
    assertEquals(true, ExcelConditionalFormattingStyleSupport.booleanProperty(implicitTrue));
    var explicitFalse =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty.Factory.newInstance();
    explicitFalse.setVal(false);
    assertEquals(false, ExcelConditionalFormattingStyleSupport.booleanProperty(explicitFalse));

    var solidWithoutForeground =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    solidWithoutForeground.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
    assertNull(
        ExcelConditionalFormattingStyleSupport.patternForegroundColor(
            solidWithoutForeground, new java.util.ArrayList<>()));
    var unsupportedForeground =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    unsupportedForeground.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
    unsupportedForeground.addNewFgColor().setTheme(4L);
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new java.util.ArrayList<>();
    assertNull(
        ExcelConditionalFormattingStyleSupport.patternForegroundColor(
            unsupportedForeground, unsupportedFeatures));
    assertEquals(
        List.of(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN), unsupportedFeatures);
  }

  @Test
  void snapshotStyleHandlesEmptyAndMalformedDifferentialPayloads() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFConditionalFormattingRule emptyRule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("A1>0");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A3")
              },
              emptyRule);
      CTCfRule emptyCtRule =
          sheet.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0);
      var emptyDxf =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf.Factory.newInstance();
      emptyDxf.addNewFont();
      emptyDxf.addNewFill();
      emptyDxf.addNewBorder();
      ExcelConditionalFormattingStyleSupport.attachStyle(
          workbook.getStylesSource(), emptyCtRule, emptyDxf);

      assertNull(
          ExcelConditionalFormattingStyleSupport.snapshotStyle(
              workbook.getStylesSource(), emptyCtRule));

      XSSFConditionalFormattingRule malformedRule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("B1>0");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("B1:B3")
              },
              malformedRule);
      CTCfRule malformedCtRule =
          sheet.getCTWorksheet().getConditionalFormattingArray(1).getCfRuleArray(0);
      var malformedDxf =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf.Factory.newInstance();
      var font = malformedDxf.addNewFont();
      font.addNewU()
          .setVal(org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues.NONE);
      font.addNewColor().setTheme(1L);
      var fill = malformedDxf.addNewFill().addNewPatternFill();
      fill.setPatternType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
      fill.addNewFgColor().setTheme(2L);
      var border = malformedDxf.addNewBorder();
      border
          .addNewTop()
          .setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
      border.getTop().addNewColor().setTheme(3L);
      ExcelConditionalFormattingStyleSupport.attachStyle(
          workbook.getStylesSource(), malformedCtRule, malformedDxf);

      ExcelDifferentialStyleSnapshot snapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(
              workbook.getStylesSource(), malformedCtRule);

      assertEquals(false, snapshot.underline());
      assertNull(snapshot.fontColor());
      assertNull(snapshot.fillColor());
      assertNull(snapshot.border());
      assertEquals(
          List.of(
              ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES,
              ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN,
              ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY),
          snapshot.unsupportedFeatures());
    }
  }

  @Test
  void applyStyleHandlesExplicitFalseUnderlineAndSparseBorders() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFConditionalFormattingRule rule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("A1>10");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:A3")
              },
              rule);
      CTCfRule ctRule = sheet.getCTWorksheet().getConditionalFormattingArray(0).getCfRuleArray(0);

      ExcelConditionalFormattingStyleSupport.applyStyle(
          workbook,
          ctRule,
          new ExcelDifferentialStyle(
              null,
              null,
              null,
              null,
              null,
              false,
              null,
              null,
              new ExcelDifferentialBorder(
                  null,
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, null),
                  null,
                  null,
                  null)));

      ExcelDifferentialStyleSnapshot snapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(workbook.getStylesSource(), ctRule);

      assertEquals(
          new ExcelDifferentialStyleSnapshot(
              null,
              null,
              null,
              null,
              null,
              false,
              null,
              null,
              new ExcelDifferentialBorder(
                  null,
                  new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, null),
                  null,
                  null,
                  null),
              List.of()),
          snapshot);
    }
  }

  @Test
  void snapshotStyleHandlesGradientFillWithoutPatternPayload() throws Exception {
    var fill =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill.Factory.parse(
            "<fill xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<gradientFill/>"
                + "</fill>");

    assertEquals(
        new ExcelConditionalFormattingStyleSupport.FillSnapshot(
            null, List.of(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN)),
        ExcelConditionalFormattingStyleSupport.snapshotFill(fill));
  }
}
