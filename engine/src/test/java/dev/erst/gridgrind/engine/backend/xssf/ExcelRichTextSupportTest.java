package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.engine.styles.ColorAdapter;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt;

/** Tests for the project-owned rich-text model and POI seam. */
class ExcelRichTextSupportTest {
  @Test
  void richTextModelsValidateInputsAndCopyRunLists() {
    assertThrows(NullPointerException.class, () -> new ResolvedRichText(null));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedRichText(List.of()));

    List<ResolvedRichTextRun> invalidRuns = new ArrayList<>();
    invalidRuns.add(null);
    assertThrows(NullPointerException.class, () -> new ResolvedRichText(invalidRuns));

    assertThrows(NullPointerException.class, () -> new ResolvedRichTextRun(null, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class, () -> new ResolvedRichTextRun("", Optional.empty()));

    List<ResolvedRichTextRun> runs =
        new ArrayList<>(
            List.of(
                new ResolvedRichTextRun("Quarterly", Optional.empty()),
                new ResolvedRichTextRun(
                    " Report",
                    Optional.of(
                        new CellFontPatch(
                            Optional.of(Boolean.TRUE),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty())))));
    ResolvedRichText richText = new ResolvedRichText(runs);
    runs.clear();

    assertEquals("Quarterly Report", richText.plainText());
    assertEquals(2, richText.runs().size());

    assertThrows(NullPointerException.class, () -> new RichTextRunReport(null, baseFont()));
    assertThrows(NullPointerException.class, () -> new RichTextRunReport("text", null));
    assertThrows(IllegalArgumentException.class, () -> new RichTextRunReport("", baseFont()));
  }

  @Test
  void textSnapshotsRequireRichTextRunsToMatchTheStoredStringValue() {
    List<RichTextRunReport> richText =
        List.copyOf(
            List.of(
                new RichTextRunReport("Quarterly", baseFont()),
                new RichTextRunReport(" Report", baseFont())));

    CellReport.TextReport snapshot =
        new CellReport.TextReport(
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            Optional.ofNullable("Quarterly Report"),
            Optional.ofNullable(defaultStyle()),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Text("Quarterly Report")),
            Optional.ofNullable(richText));

    assertEquals("TEXT", snapshot.type());
    assertEquals(richText, snapshot.runs().orElseThrow());

    IllegalArgumentException mismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new CellReport.TextReport(
                    new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                    Optional.ofNullable("Quarterly Report"),
                    Optional.ofNullable(defaultStyle()),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(
                        new dev.erst.gridgrind.core.model.value.CellValue.Text("Different")),
                    Optional.ofNullable(richText)));
    assertEquals("runs text must concatenate to the value", mismatch.getMessage());
  }

  @Test
  void snapshotReturnsNullForScalarStringsAndUsesTheBaseFontWhenRunsCarryNoPatch()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertThrows(
          NullPointerException.class,
          () ->
              StyleAdapter.toPoiRichText(
                  null,
                  new ResolvedRichText(
                      List.of(new ResolvedRichTextRun("text", Optional.empty())))));
      assertThrows(NullPointerException.class, () -> StyleAdapter.toPoiRichText(workbook, null));
      assertThrows(
          NullPointerException.class,
          () -> StyleAdapter.snapshot(null, new XSSFRichTextString("plain"), baseFont()));
      assertThrows(
          NullPointerException.class, () -> StyleAdapter.snapshot(workbook, null, baseFont()));
      assertThrows(
          NullPointerException.class,
          () -> StyleAdapter.snapshot(workbook, new XSSFRichTextString("plain"), null));

      assertEquals(
          Optional.empty(),
          StyleAdapter.snapshot(workbook, new XSSFRichTextString("plain"), baseFont()));

      XSSFRichTextString runWithoutProperties = new XSSFRichTextString();
      runWithoutProperties.getCTRst().addNewR().setT("Base");
      List<RichTextRunReport> inheritedWithoutProperties =
          StyleAdapter.snapshot(workbook, runWithoutProperties, baseFont()).orElseThrow();

      assertNotNull(inheritedWithoutProperties);
      assertThrows(
          UnsupportedOperationException.class,
          () -> inheritedWithoutProperties.add(new RichTextRunReport("extra", baseFont())));
      assertEquals(
          "Base",
          inheritedWithoutProperties.stream()
              .map(RichTextRunReport::text)
              .collect(java.util.stream.Collectors.joining()));
      assertEquals(baseFont(), inheritedWithoutProperties.get(0).font());

      XSSFRichTextString runWithEmptyProperties = new XSSFRichTextString();
      CTRElt run = runWithEmptyProperties.getCTRst().addNewR();
      run.setT(" Still Base");
      run.addNewRPr();
      List<RichTextRunReport> inheritedWithEmptyProperties =
          StyleAdapter.snapshot(workbook, runWithEmptyProperties, baseFont()).orElseThrow();

      assertNotNull(inheritedWithEmptyProperties);
      assertEquals(
          " Still Base",
          inheritedWithEmptyProperties.stream()
              .map(RichTextRunReport::text)
              .collect(java.util.stream.Collectors.joining()));
      assertEquals(baseFont(), inheritedWithEmptyProperties.get(0).font());
    }
  }

  @Test
  void richTextRoundTripsThemeAndIndexedFontColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFRichTextString richText =
          StyleAdapter.toPoiRichText(
              workbook,
              new ResolvedRichText(
                  List.of(
                      new ResolvedRichTextRun(
                          "Lead",
                          Optional.of(
                              new CellFontPatch(
                                  Optional.of(true),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.of(ColorSpec.theme(4, -0.20d)),
                                  Optional.empty(),
                                  Optional.empty()))),
                      new ResolvedRichTextRun(" ", Optional.empty()),
                      new ResolvedRichTextRun(
                          "review scheduled",
                          Optional.of(
                              new CellFontPatch(
                                  Optional.empty(),
                                  Optional.of(true),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.of(
                                      ColorSpec.indexed(
                                          Short.toUnsignedInt(
                                              IndexedColors.DARK_GREEN.getIndex()))),
                                  Optional.empty(),
                                  Optional.empty()))))));

      List<RichTextRunReport> snapshot =
          StyleAdapter.snapshot(workbook, richText, baseFont()).orElseThrow();

      assertEquals(CellColorReport.theme(4, -0.20d), snapshot.get(0).font().fontColor());
      assertEquals(baseFont().fontColor(), snapshot.get(1).font().fontColor());
      assertEquals(
          CellColorReport.indexed(Short.toUnsignedInt(IndexedColors.DARK_GREEN.getIndex())),
          snapshot.get(2).font().fontColor());
    }
  }

  @Test
  void roundTripsExplicitRunOverridesAndInheritedBaseFontValues() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ResolvedRichText richText =
          new ResolvedRichText(
              List.of(
                  new ResolvedRichTextRun(
                      "Alpha",
                      Optional.of(
                          new CellFontPatch(
                              Optional.of(Boolean.FALSE),
                              Optional.of(Boolean.FALSE),
                              Optional.of("Courier New"),
                              Optional.of(new FontHeightSpec.Points(new BigDecimal("14"))),
                              Optional.of(ColorSpec.rgb(new RgbHex("#123456"))),
                              Optional.of(Boolean.FALSE),
                              Optional.of(Boolean.TRUE)))),
                  new ResolvedRichTextRun(
                      " Beta",
                      Optional.of(
                          new CellFontPatch(
                              Optional.of(Boolean.TRUE),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of(ColorSpec.rgb(new RgbHex("#ABCDEF"))),
                              Optional.of(Boolean.TRUE),
                              Optional.of(Boolean.FALSE)))),
                  new ResolvedRichTextRun(
                      " Delta",
                      Optional.of(
                          new CellFontPatch(
                              Optional.empty(),
                              Optional.of(Boolean.FALSE),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))),
                  new ResolvedRichTextRun(" Gamma", Optional.empty())));

      List<RichTextRunReport> snapshot =
          StyleAdapter.snapshot(
                  workbook, StyleAdapter.toPoiRichText(workbook, richText), baseFont())
              .orElseThrow();

      assertNotNull(snapshot);
      assertEquals(
          "Alpha Beta Delta Gamma",
          snapshot.stream()
              .map(RichTextRunReport::text)
              .collect(java.util.stream.Collectors.joining()));
      assertEquals(4, snapshot.size());

      RichTextRunReport alpha = snapshot.get(0);
      assertEquals("Alpha", alpha.text());
      assertFalse(alpha.font().bold());
      assertFalse(alpha.font().italic());
      assertEquals("Courier New", alpha.font().fontName());
      assertEquals(FontHeightReport.fromPoints(new BigDecimal("14")), alpha.font().fontHeight());
      assertEquals(CellColorReport.rgb(new RgbHex("#123456")), alpha.font().fontColor());
      assertFalse(alpha.font().underline());
      assertTrue(alpha.font().strikeout());

      RichTextRunReport beta = snapshot.get(1);
      assertEquals(" Beta", beta.text());
      assertTrue(beta.font().bold());
      assertTrue(beta.font().italic());
      assertEquals(baseFont().fontName(), beta.font().fontName());
      assertEquals(baseFont().fontHeight(), beta.font().fontHeight());
      assertEquals(CellColorReport.rgb(new RgbHex("#ABCDEF")), beta.font().fontColor());
      assertTrue(beta.font().underline());
      assertFalse(beta.font().strikeout());

      RichTextRunReport delta = snapshot.get(2);
      assertEquals(" Delta", delta.text());
      assertEquals(baseFont().bold(), delta.font().bold());
      assertFalse(delta.font().italic());
      assertEquals(baseFont().fontName(), delta.font().fontName());
      assertEquals(baseFont().fontHeight(), delta.font().fontHeight());
      assertEquals(baseFont().fontColor(), delta.font().fontColor());
      assertEquals(baseFont().underline(), delta.font().underline());
      assertEquals(baseFont().strikeout(), delta.font().strikeout());

      RichTextRunReport gamma = snapshot.get(3);
      assertEquals(" Gamma", gamma.text());
      assertEquals(baseFont(), gamma.font());
    }
  }

  @Test
  void snapshotTreatsImplicitBooleanAndUnderlineRunPropertiesAsTrue() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.getStylesSource().ensureThemesTable();

      XSSFRichTextString richText = new XSSFRichTextString();
      CTRElt run = richText.getCTRst().addNewR();
      run.setT("Flags");
      CTRPrElt properties = run.addNewRPr();
      properties.addNewB();
      properties.addNewI();
      properties.addNewStrike();
      properties.addNewU();
      properties.addNewRFont().setVal("Fira Code");
      properties.addNewSz().setVal(18.0d);
      properties
          .addNewColor()
          .setRgb(
              ColorAdapter.toXssfColor(
                      workbook, new dev.erst.gridgrind.core.vocabulary.RgbHex("#00AA11"))
                  .getARGB());

      List<RichTextRunReport> snapshot =
          StyleAdapter.snapshot(workbook, richText, baseFont()).orElseThrow();

      assertNotNull(snapshot);
      assertEquals(
          "Flags",
          snapshot.stream()
              .map(RichTextRunReport::text)
              .collect(java.util.stream.Collectors.joining()));
      assertEquals(1, snapshot.size());
      CellFontReport font = snapshot.get(0).font();
      assertTrue(font.bold());
      assertTrue(font.italic());
      assertEquals("Fira Code", font.fontName());
      assertEquals(FontHeightReport.fromPoints(new BigDecimal("18")), font.fontHeight());
      assertEquals(CellColorReport.rgb(new RgbHex("#00AA11")), font.fontColor());
      assertTrue(font.underline());
      assertTrue(font.strikeout());
    }
  }

  @Test
  void snapshotRetainsRunsWhoseOnlyOverridesAppearLateInTheFontPatchChain() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ResolvedRichText richText =
          new ResolvedRichText(
              List.of(
                  new ResolvedRichTextRun(
                      "Name",
                      Optional.of(
                          new CellFontPatch(
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of("Courier New"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))),
                  new ResolvedRichTextRun(
                      " Size",
                      Optional.of(
                          new CellFontPatch(
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of(new FontHeightSpec.Points(new BigDecimal("13"))),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))),
                  new ResolvedRichTextRun(
                      " Color",
                      Optional.of(
                          new CellFontPatch(
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of(ColorSpec.rgb(new RgbHex("#ABCDEF"))),
                              Optional.empty(),
                              Optional.empty()))),
                  new ResolvedRichTextRun(
                      " Underline",
                      Optional.of(
                          new CellFontPatch(
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of(Boolean.FALSE),
                              Optional.empty()))),
                  new ResolvedRichTextRun(
                      " Strike",
                      Optional.of(
                          new CellFontPatch(
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.of(Boolean.TRUE))))));

      List<RichTextRunReport> snapshot =
          StyleAdapter.snapshot(
                  workbook, StyleAdapter.toPoiRichText(workbook, richText), baseFont())
              .orElseThrow();

      assertNotNull(snapshot);
      assertEquals(5, snapshot.size());
      assertEquals("Courier New", snapshot.get(0).font().fontName());
      assertEquals(
          FontHeightReport.fromPoints(new BigDecimal("13")), snapshot.get(1).font().fontHeight());
      assertEquals(CellColorReport.rgb(new RgbHex("#ABCDEF")), snapshot.get(2).font().fontColor());
      assertFalse(snapshot.get(3).font().underline());
      assertTrue(snapshot.get(4).font().strikeout());
    }
  }

  private static CellFontReport baseFont() {
    return new CellFontReport(
        true,
        true,
        "Aptos",
        FontHeightReport.fromPoints(new BigDecimal("11")),
        CellColorReport.rgb(new RgbHex("#654321")),
        true,
        false);
  }

  private static CellStyleReport defaultStyle() {
    return new CellStyleReport(
        "",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        baseFont(),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
  }
}
