package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
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
              Optional.of("0.00"),
              Optional.of(true),
              Optional.of(false),
              Optional.ofNullable(ExcelFontHeight.fromPoints(BigDecimal.valueOf(11))),
              Optional.of(ExcelColor.theme(3, Optional.of(0.25d))),
              Optional.of(true),
              Optional.of(true),
              Optional.of(ExcelColor.indexed(12, Optional.empty())),
              Optional.ofNullable(
                  new ExcelDifferentialBorder(
                      new ExcelBorderSide(
                          Optional.of(ExcelBorderStyle.THIN),
                          Optional.of(ExcelColor.theme(5, Optional.of(-0.2d)))),
                      null,
                      null,
                      null,
                      null)));

      ExcelConditionalFormattingStyleSupport.applyStyle(workbook, ctRule, authoredStyle);

      ExcelDifferentialStyleSnapshot snapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(workbook.getStylesSource(), ctRule);

      assertEquals(
          new ExcelDifferentialStyleSnapshot(
              "0.00",
              true,
              false,
              ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
              ExcelColor.theme(3, Optional.of(0.25d)),
              true,
              true,
              ExcelColor.indexed(12, Optional.empty()),
              new ExcelDifferentialBorder(
                  null,
                  new ExcelBorderSide(
                      Optional.of(ExcelBorderStyle.THIN),
                      Optional.of(ExcelColor.theme(5, Optional.of(-0.2d)))),
                  new ExcelBorderSide(
                      Optional.of(ExcelBorderStyle.THIN),
                      Optional.of(ExcelColor.theme(5, Optional.of(-0.2d)))),
                  new ExcelBorderSide(
                      Optional.of(ExcelBorderStyle.THIN),
                      Optional.of(ExcelColor.theme(5, Optional.of(-0.2d)))),
                  new ExcelBorderSide(
                      Optional.of(ExcelBorderStyle.THIN),
                      Optional.of(ExcelColor.theme(5, Optional.of(-0.2d))))),
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
      assertEquals(ExcelColor.theme(1), snapshot.fontColor());
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
  void colorHelpersPreserveEverySupportedRawColorReference() {
    var argbColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    argbColor.setRgb(new byte[] {(byte) 0xFF, 0x10, 0x20, 0x30});
    var rgbColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    rgbColor.setRgb(new byte[] {0x40, 0x50, 0x60});
    var unsupportedColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    unsupportedColor.setRgb(new byte[] {0x01, 0x02});
    var themeColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    themeColor.setTheme(5L);
    themeColor.setTint(-0.2d);
    var indexedColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
    indexedColor.setIndexed(12L);
    var absentColor =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();

    assertEquals(Optional.empty(), ExcelConditionalFormattingColorSupport.optionalColor(null));
    assertEquals(
        Optional.of(ExcelColor.rgb("#102030")),
        ExcelConditionalFormattingColorSupport.optionalColor(argbColor));
    assertEquals(
        Optional.of(ExcelColor.rgb("#405060")),
        ExcelConditionalFormattingColorSupport.optionalColor(rgbColor));
    assertEquals(
        Optional.empty(), ExcelConditionalFormattingColorSupport.optionalColor(unsupportedColor));
    assertEquals(
        Optional.of(ExcelColor.theme(5, Optional.of(-0.2d))),
        ExcelConditionalFormattingColorSupport.optionalColor(themeColor));
    assertEquals(
        Optional.of(ExcelColor.indexed(12)),
        ExcelConditionalFormattingColorSupport.optionalColor(indexedColor));
    assertEquals(
        Optional.empty(), ExcelConditionalFormattingColorSupport.optionalColor(absentColor));
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
    var themedSide =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
    themedSide.setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
    themedSide.addNewColor().setTheme(1L);

    assertNull(ExcelConditionalFormattingStyleSupport.snapshotBorderSide(null));
    assertNull(ExcelConditionalFormattingStyleSupport.snapshotBorderSide(emptySide));
    assertEquals(
        new ExcelBorderSide(
            Optional.of(ExcelBorderStyle.DASH_DOT), Optional.of(ExcelColor.rgb("#012345"))),
        ExcelConditionalFormattingStyleSupport.snapshotBorderSide(coloredSide));
    assertEquals(
        new ExcelBorderSide(Optional.of(ExcelBorderStyle.THIN), Optional.of(ExcelColor.theme(1))),
        ExcelConditionalFormattingStyleSupport.snapshotBorderSide(themedSide));

    var colorOnlySide =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
    colorOnlySide.addNewColor().setRgb(new byte[] {(byte) 0xFF, 0x22, 0x33, 0x44});
    assertEquals(
        new ExcelBorderSide(Optional.empty(), Optional.of(ExcelColor.rgb("#223344"))),
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

    ExcelBorderSide thinSide = new ExcelBorderSide(ExcelBorderStyle.THIN);

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
    var malformedForeground =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    malformedForeground.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
    malformedForeground.addNewFgColor().setRgb(new byte[] {0x01, 0x02});
    List<ExcelConditionalFormattingUnsupportedFeature> malformedFeatures =
        new java.util.ArrayList<>();
    assertNull(
        ExcelConditionalFormattingStyleSupport.patternForegroundColor(
            malformedForeground, malformedFeatures));
    assertEquals(
        List.of(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN), malformedFeatures);
    var themedForeground =
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill.Factory.newInstance();
    themedForeground.setPatternType(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
    themedForeground.addNewFgColor().setTheme(4L);
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new java.util.ArrayList<>();
    assertEquals(
        ExcelColor.theme(4),
        ExcelConditionalFormattingStyleSupport.patternForegroundColor(
            themedForeground, unsupportedFeatures));
    assertEquals(List.of(), unsupportedFeatures);
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
      assertEquals(ExcelColor.theme(1), snapshot.fontColor());
      assertEquals(ExcelColor.theme(2), snapshot.fillColor());
      assertEquals(
          new ExcelDifferentialBorder(
              null,
              new ExcelBorderSide(
                  Optional.of(ExcelBorderStyle.THIN), Optional.of(ExcelColor.theme(3))),
              null,
              null,
              null),
          snapshot.border());
      assertEquals(List.of(), snapshot.unsupportedFeatures());

      XSSFConditionalFormattingRule unsupportedRule =
          sheet.getSheetConditionalFormatting().createConditionalFormattingRule("C1>0");
      sheet
          .getSheetConditionalFormatting()
          .addConditionalFormatting(
              new org.apache.poi.ss.util.CellRangeAddress[] {
                org.apache.poi.ss.util.CellRangeAddress.valueOf("C1:C3")
              },
              unsupportedRule);
      CTCfRule unsupportedCtRule =
          sheet.getCTWorksheet().getConditionalFormattingArray(2).getCfRuleArray(0);
      var unsupportedDxf =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf.Factory.newInstance();
      unsupportedDxf.addNewFont().addNewColor().setRgb(new byte[] {0x01, 0x02});
      unsupportedDxf.addNewBorder().addNewTop().addNewColor().setRgb(new byte[] {0x01, 0x02});
      ExcelConditionalFormattingStyleSupport.attachStyle(
          workbook.getStylesSource(), unsupportedCtRule, unsupportedDxf);

      ExcelDifferentialStyleSnapshot unsupportedSnapshot =
          ExcelConditionalFormattingStyleSupport.snapshotStyle(
              workbook.getStylesSource(), unsupportedCtRule);

      assertNull(unsupportedSnapshot.fontColor());
      assertNull(unsupportedSnapshot.border());
      assertEquals(
          List.of(
              ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES,
              ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY),
          unsupportedSnapshot.unsupportedFeatures());
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
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.empty(),
              Optional.of(false),
              Optional.empty(),
              Optional.empty(),
              Optional.ofNullable(
                  new ExcelDifferentialBorder(
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.THIN),
                      new ExcelBorderSide(Optional.empty(), Optional.of(ExcelColor.indexed(9))),
                      null,
                      null))));

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
                  new ExcelBorderSide(ExcelBorderStyle.THIN),
                  new ExcelBorderSide(Optional.empty(), Optional.of(ExcelColor.indexed(9))),
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
