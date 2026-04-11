package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
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
    assertThrows(NullPointerException.class, () -> new ExcelRichText(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRichText(List.of()));

    List<ExcelRichTextRun> invalidRuns = new ArrayList<>();
    invalidRuns.add(null);
    assertThrows(NullPointerException.class, () -> new ExcelRichText(invalidRuns));

    assertThrows(NullPointerException.class, () -> new ExcelRichTextRun(null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRichTextRun("", null));

    List<ExcelRichTextRun> runs =
        new ArrayList<>(
            List.of(
                new ExcelRichTextRun("Quarterly", null),
                new ExcelRichTextRun(
                    " Report",
                    new ExcelCellFont(Boolean.TRUE, null, null, null, null, null, null))));
    ExcelRichText richText = new ExcelRichText(runs);
    runs.clear();

    assertEquals("Quarterly Report", richText.plainText());
    assertEquals(2, richText.runs().size());

    assertThrows(NullPointerException.class, () -> new ExcelRichTextSnapshot(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRichTextSnapshot(List.of()));

    List<ExcelRichTextRunSnapshot> invalidSnapshots = new ArrayList<>();
    invalidSnapshots.add(null);
    assertThrows(NullPointerException.class, () -> new ExcelRichTextSnapshot(invalidSnapshots));

    assertThrows(NullPointerException.class, () -> new ExcelRichTextRunSnapshot(null, baseFont()));
    assertThrows(NullPointerException.class, () -> new ExcelRichTextRunSnapshot("text", null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelRichTextRunSnapshot("", baseFont()));

    List<ExcelRichTextRunSnapshot> snapshotRuns =
        new ArrayList<>(
            List.of(
                new ExcelRichTextRunSnapshot("Quarterly", baseFont()),
                new ExcelRichTextRunSnapshot(" Report", baseFont())));
    ExcelRichTextSnapshot snapshot = new ExcelRichTextSnapshot(snapshotRuns);
    snapshotRuns.clear();

    assertEquals("Quarterly Report", snapshot.plainText());
    assertEquals(2, snapshot.runs().size());
  }

  @Test
  void textSnapshotsRequireRichTextRunsToMatchTheStoredStringValue() {
    ExcelRichTextSnapshot richText =
        new ExcelRichTextSnapshot(
            List.of(
                new ExcelRichTextRunSnapshot("Quarterly", baseFont()),
                new ExcelRichTextRunSnapshot(" Report", baseFont())));

    ExcelCellSnapshot.TextSnapshot snapshot =
        new ExcelCellSnapshot.TextSnapshot(
            "A1",
            "STRING",
            "Quarterly Report",
            defaultStyle(),
            ExcelCellMetadataSnapshot.empty(),
            "Quarterly Report",
            richText);

    assertEquals("STRING", snapshot.effectiveType());
    assertEquals(richText, snapshot.richText());

    IllegalArgumentException mismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ExcelCellSnapshot.TextSnapshot(
                    "A1",
                    "STRING",
                    "Quarterly Report",
                    defaultStyle(),
                    ExcelCellMetadataSnapshot.empty(),
                    "Different",
                    richText));
    assertEquals("richText run text must concatenate to the stringValue", mismatch.getMessage());
  }

  @Test
  void snapshotReturnsNullForScalarStringsAndUsesTheBaseFontWhenRunsCarryNoPatch()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertThrows(
          NullPointerException.class,
          () ->
              ExcelRichTextSupport.toPoiRichText(
                  null, new ExcelRichText(List.of(new ExcelRichTextRun("text", null)))));
      assertThrows(
          NullPointerException.class, () -> ExcelRichTextSupport.toPoiRichText(workbook, null));
      assertThrows(
          NullPointerException.class,
          () -> ExcelRichTextSupport.snapshot(null, new XSSFRichTextString("plain"), baseFont()));
      assertThrows(
          NullPointerException.class,
          () -> ExcelRichTextSupport.snapshot(workbook, null, baseFont()));
      assertThrows(
          NullPointerException.class,
          () -> ExcelRichTextSupport.snapshot(workbook, new XSSFRichTextString("plain"), null));

      assertNull(
          ExcelRichTextSupport.snapshot(workbook, new XSSFRichTextString("plain"), baseFont()));

      XSSFRichTextString runWithoutProperties = new XSSFRichTextString();
      runWithoutProperties.getCTRst().addNewR().setT("Base");
      ExcelRichTextSnapshot inheritedWithoutProperties =
          ExcelRichTextSupport.snapshot(workbook, runWithoutProperties, baseFont());

      assertNotNull(inheritedWithoutProperties);
      assertEquals("Base", inheritedWithoutProperties.plainText());
      assertEquals(baseFont(), inheritedWithoutProperties.runs().get(0).font());

      XSSFRichTextString runWithEmptyProperties = new XSSFRichTextString();
      CTRElt run = runWithEmptyProperties.getCTRst().addNewR();
      run.setT(" Still Base");
      run.addNewRPr();
      ExcelRichTextSnapshot inheritedWithEmptyProperties =
          ExcelRichTextSupport.snapshot(workbook, runWithEmptyProperties, baseFont());

      assertNotNull(inheritedWithEmptyProperties);
      assertEquals(" Still Base", inheritedWithEmptyProperties.plainText());
      assertEquals(baseFont(), inheritedWithEmptyProperties.runs().get(0).font());
    }
  }

  @Test
  void richTextRoundTripsThemeAndIndexedFontColors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFRichTextString richText =
          ExcelRichTextSupport.toPoiRichText(
              workbook,
              new ExcelRichText(
                  List.of(
                      new ExcelRichTextRun(
                          "Lead",
                          new ExcelCellFont(
                              true,
                              null,
                              null,
                              null,
                              new ExcelColor(null, 4, null, -0.20d),
                              null,
                              null)),
                      new ExcelRichTextRun(" ", null),
                      new ExcelRichTextRun(
                          "review scheduled",
                          new ExcelCellFont(
                              null,
                              true,
                              null,
                              null,
                              new ExcelColor(
                                  null,
                                  null,
                                  Short.toUnsignedInt(IndexedColors.DARK_GREEN.getIndex()),
                                  null),
                              null,
                              null)))));

      ExcelRichTextSnapshot snapshot =
          ExcelRichTextSupport.snapshot(workbook, richText, baseFont());

      assertEquals(
          new ExcelColorSnapshot(null, 4, null, -0.20d), snapshot.runs().get(0).font().fontColor());
      assertEquals(baseFont().fontColor(), snapshot.runs().get(1).font().fontColor());
      assertEquals(
          new ExcelColorSnapshot(
              null, null, Short.toUnsignedInt(IndexedColors.DARK_GREEN.getIndex()), null),
          snapshot.runs().get(2).font().fontColor());
    }
  }

  @Test
  void roundTripsExplicitRunOverridesAndInheritedBaseFontValues() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ExcelRichText richText =
          new ExcelRichText(
              List.of(
                  new ExcelRichTextRun(
                      "Alpha",
                      new ExcelCellFont(
                          Boolean.FALSE,
                          Boolean.FALSE,
                          "Courier New",
                          ExcelFontHeight.fromPoints(new BigDecimal("14")),
                          new ExcelColor("#123456"),
                          Boolean.FALSE,
                          Boolean.TRUE)),
                  new ExcelRichTextRun(
                      " Beta",
                      new ExcelCellFont(
                          Boolean.TRUE,
                          null,
                          null,
                          null,
                          new ExcelColor("#ABCDEF"),
                          Boolean.TRUE,
                          Boolean.FALSE)),
                  new ExcelRichTextRun(
                      " Delta",
                      new ExcelCellFont(null, Boolean.FALSE, null, null, null, null, null)),
                  new ExcelRichTextRun(" Gamma", null)));

      ExcelRichTextSnapshot snapshot =
          ExcelRichTextSupport.snapshot(
              workbook, ExcelRichTextSupport.toPoiRichText(workbook, richText), baseFont());

      assertNotNull(snapshot);
      assertEquals("Alpha Beta Delta Gamma", snapshot.plainText());
      assertEquals(4, snapshot.runs().size());

      ExcelRichTextRunSnapshot alpha = snapshot.runs().get(0);
      assertEquals("Alpha", alpha.text());
      assertFalse(alpha.font().bold());
      assertFalse(alpha.font().italic());
      assertEquals("Courier New", alpha.font().fontName());
      assertEquals(ExcelFontHeight.fromPoints(new BigDecimal("14")), alpha.font().fontHeight());
      assertEquals(new ExcelColorSnapshot("#123456"), alpha.font().fontColor());
      assertFalse(alpha.font().underline());
      assertTrue(alpha.font().strikeout());

      ExcelRichTextRunSnapshot beta = snapshot.runs().get(1);
      assertEquals(" Beta", beta.text());
      assertTrue(beta.font().bold());
      assertTrue(beta.font().italic());
      assertEquals(baseFont().fontName(), beta.font().fontName());
      assertEquals(baseFont().fontHeight(), beta.font().fontHeight());
      assertEquals(new ExcelColorSnapshot("#ABCDEF"), beta.font().fontColor());
      assertTrue(beta.font().underline());
      assertFalse(beta.font().strikeout());

      ExcelRichTextRunSnapshot delta = snapshot.runs().get(2);
      assertEquals(" Delta", delta.text());
      assertEquals(baseFont().bold(), delta.font().bold());
      assertFalse(delta.font().italic());
      assertEquals(baseFont().fontName(), delta.font().fontName());
      assertEquals(baseFont().fontHeight(), delta.font().fontHeight());
      assertEquals(baseFont().fontColor(), delta.font().fontColor());
      assertEquals(baseFont().underline(), delta.font().underline());
      assertEquals(baseFont().strikeout(), delta.font().strikeout());

      ExcelRichTextRunSnapshot gamma = snapshot.runs().get(3);
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
          .setRgb(ExcelRgbColorSupport.toXssfColor(workbook, "#00AA11").getARGB());

      ExcelRichTextSnapshot snapshot =
          ExcelRichTextSupport.snapshot(workbook, richText, baseFont());

      assertNotNull(snapshot);
      assertEquals("Flags", snapshot.plainText());
      assertEquals(1, snapshot.runs().size());
      ExcelCellFontSnapshot font = snapshot.runs().get(0).font();
      assertTrue(font.bold());
      assertTrue(font.italic());
      assertEquals("Fira Code", font.fontName());
      assertEquals(ExcelFontHeight.fromPoints(new BigDecimal("18")), font.fontHeight());
      assertEquals(new ExcelColorSnapshot("#00AA11"), font.fontColor());
      assertTrue(font.underline());
      assertTrue(font.strikeout());
    }
  }

  @Test
  void snapshotRetainsRunsWhoseOnlyOverridesAppearLateInTheFontPatchChain() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ExcelRichText richText =
          new ExcelRichText(
              List.of(
                  new ExcelRichTextRun(
                      "Name", new ExcelCellFont(null, null, "Courier New", null, null, null, null)),
                  new ExcelRichTextRun(
                      " Size",
                      new ExcelCellFont(
                          null,
                          null,
                          null,
                          ExcelFontHeight.fromPoints(new BigDecimal("13")),
                          null,
                          null,
                          null)),
                  new ExcelRichTextRun(
                      " Color",
                      new ExcelCellFont(
                          null, null, null, null, new ExcelColor("#ABCDEF"), null, null)),
                  new ExcelRichTextRun(
                      " Underline",
                      new ExcelCellFont(null, null, null, null, null, Boolean.FALSE, null)),
                  new ExcelRichTextRun(
                      " Strike",
                      new ExcelCellFont(null, null, null, null, null, null, Boolean.TRUE))));

      ExcelRichTextSnapshot snapshot =
          ExcelRichTextSupport.snapshot(
              workbook, ExcelRichTextSupport.toPoiRichText(workbook, richText), baseFont());

      assertNotNull(snapshot);
      assertEquals(5, snapshot.runs().size());
      assertEquals("Courier New", snapshot.runs().get(0).font().fontName());
      assertEquals(
          ExcelFontHeight.fromPoints(new BigDecimal("13")),
          snapshot.runs().get(1).font().fontHeight());
      assertEquals(new ExcelColorSnapshot("#ABCDEF"), snapshot.runs().get(2).font().fontColor());
      assertFalse(snapshot.runs().get(3).font().underline());
      assertTrue(snapshot.runs().get(4).font().strikeout());
    }
  }

  private static ExcelCellFontSnapshot baseFont() {
    return new ExcelCellFontSnapshot(
        true,
        true,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        new ExcelColorSnapshot("#654321"),
        true,
        false);
  }

  private static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        baseFont(),
        new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, null),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }
}
