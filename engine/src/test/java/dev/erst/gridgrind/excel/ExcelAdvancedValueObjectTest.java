package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused contract tests for advanced workbook-core value objects added during parity work. */
class ExcelAdvancedValueObjectTest {
  @Test
  void autofilterValueObjectsNormalizeAndValidate() {
    var customCondition = new ExcelAutofilterFilterCriterion.CustomCondition("equal", "Ada");
    var values =
        new ExcelAutofilterFilterCriterion.Values(new ArrayList<>(List.of("Queued", "")), true);
    var custom = new ExcelAutofilterFilterCriterion.Custom(true, List.of(customCondition));
    var dynamic = new ExcelAutofilterFilterCriterion.Dynamic("today", 1.0d, 2.0d);
    var top10 = new ExcelAutofilterFilterCriterion.Top10(10, false, true);
    var color =
        new ExcelAutofilterFilterCriterion.Color(false, new ExcelColor(null, null, 10, -0.25d));
    var icon = new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2);
    var column = new ExcelAutofilterFilterColumn(4L, false, color);
    var sortCondition =
        new ExcelAutofilterSortCondition("A2:A5", true, null, new ExcelColor("#AABBCC"), 1);
    var sortState = new ExcelAutofilterSortState("A1:F5", true, true, "", List.of(sortCondition));

    assertEquals(List.of("Queued", ""), values.values());
    assertEquals(List.of(customCondition), custom.conditions());
    assertEquals("today", dynamic.type());
    assertEquals(10, top10.value());
    assertEquals(new ExcelColor(null, null, 10, -0.25d), color.color());
    assertEquals("3TrafficLights1", icon.iconSet());
    assertEquals(4L, column.columnId());
    assertEquals("", sortCondition.sortBy());
    assertEquals(List.of(sortCondition), sortState.conditions());

    List<String> valuesWithNull = new ArrayList<>(Arrays.asList("Queued", null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterFilterCriterion.Values(valuesWithNull, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Custom(false, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Dynamic(" ", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Dynamic("today", Double.NaN, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Top10(0, true, false));
    assertThrows(
        NullPointerException.class, () -> new ExcelAutofilterFilterCriterion.Color(true, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelAutofilterFilterCriterion.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.CustomCondition(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterion.CustomCondition("equal", " "));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelAutofilterFilterColumn(-1L, true, values));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterSortCondition(null, false, "", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortCondition(" ", false, "", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortCondition("A1", false, "", null, -1));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelAutofilterSortState("A1:B2", false, false, null, List.of(sortCondition)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortState(" ", false, false, "", List.of(sortCondition)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortState("A1:B2", false, false, "", List.of()));
  }

  @Test
  void styleValueObjectsNormalizeDefaultsAndValidate() {
    var themeColor = new ExcelColor(null, 6, null, -0.35d);
    var indexedColor = new ExcelColor(null, null, 10, null);
    var gradient =
        new ExcelGradientFill(
            " PATH ",
            null,
            0.1d,
            0.2d,
            0.3d,
            0.4d,
            List.of(
                new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                new ExcelGradientStop(1.0d, indexedColor)));
    var gradientFill = new ExcelCellFill(null, null, null, gradient);
    var printSetup =
        new ExcelPrintSetup(
            new ExcelPrintMargins(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d),
            true,
            true,
            8,
            true,
            true,
            2,
            true,
            4,
            List.of(1, 3),
            List.of(2));
    var printLayout =
        new ExcelPrintLayout(
            new ExcelPrintLayout.Area.Range("A1:C3"),
            ExcelPrintOrientation.LANDSCAPE,
            new ExcelPrintLayout.Scaling.Fit(1, 0),
            new ExcelPrintLayout.TitleRows.Band(0, 0),
            new ExcelPrintLayout.TitleColumns.Band(0, 1),
            new ExcelHeaderFooterText("Left", "Center", "Right"),
            new ExcelHeaderFooterText("Footer", "", ""));

    assertEquals("#AA00CC", new ExcelColor("#aa00cc").rgb());
    assertEquals(6, themeColor.theme());
    assertEquals(10, indexedColor.indexed());
    assertEquals("PATH", gradient.type());
    assertEquals(2, gradient.stops().size());
    assertSame(gradient, gradientFill.gradient());
    assertEquals(8, printSetup.paperSize());
    assertEquals(ExcelPrintSetup.defaults(), printLayout.setup());
    assertEquals(new ExcelPrintLayout.Area.None(), ExcelPrintLayout.defaults().printArea());

    assertThrows(IllegalArgumentException.class, () -> new ExcelColor(" "));
    assertThrows(IllegalArgumentException.class, () -> new ExcelColor("#12345"));
    assertThrows(IllegalArgumentException.class, () -> new ExcelColor(null, -1, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelColor(null, null, -1, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelColor(null, null, null, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelColor(null, 1, null, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelGradientFill(" ", null, null, null, null, null, List.of()));
    assertEquals(
        "LINEAR",
        new ExcelGradientFill(
                null,
                null,
                null,
                null,
                null,
                null,
                List.of(
                    new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                    new ExcelGradientStop(1.0d, new ExcelColor("#445566"))))
            .type());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelGradientFill(
                "LINEAR",
                Double.POSITIVE_INFINITY,
                null,
                null,
                null,
                null,
                List.of(
                    new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                    new ExcelGradientStop(1.0d, new ExcelColor("#445566")))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelGradientFill(
                "radial",
                null,
                null,
                null,
                null,
                null,
                List.of(
                    new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                    new ExcelGradientStop(1.0d, new ExcelColor("#445566")))));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelGradientFill("LINEAR", null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelGradientFill(
                "LINEAR",
                null,
                null,
                null,
                null,
                null,
                List.of(new ExcelGradientStop(0.0d, new ExcelColor("#112233")))));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelGradientFill(
                "LINEAR",
                null,
                null,
                null,
                null,
                null,
                Arrays.asList(new ExcelGradientStop(0.0d, new ExcelColor("#112233")), null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelGradientStop(-0.1d, new ExcelColor("#112233")));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCellFill(null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFill(null, null, new ExcelColor("#112233"), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellFill(
                ExcelFillPattern.SOLID,
                new ExcelColor("#112233"),
                null,
                new ExcelGradientFill(
                    "LINEAR",
                    null,
                    null,
                    null,
                    null,
                    null,
                    List.of(
                        new ExcelGradientStop(0.0d, new ExcelColor("#112233")),
                        new ExcelGradientStop(1.0d, new ExcelColor("#445566"))))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFill(ExcelFillPattern.NONE, new ExcelColor("#112233"), null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellFill(
                ExcelFillPattern.SOLID,
                new ExcelColor("#112233"),
                new ExcelColor("#445566"),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                -1,
                false,
                false,
                1,
                false,
                1,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                1,
                false,
                false,
                -1,
                false,
                1,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                -1,
                List.of(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetup(
                printSetup.margins(),
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                List.of(-1),
                List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPrintMargins(0, 0, Double.NaN, 0, 0, 0));
  }

  @Test
  void namedRangeCommentAndTableContractsSupportAdvancedForms() {
    var authoringComment =
        new ExcelComment(
            "Quarterly Review",
            "GridGrind",
            true,
            new ExcelRichText(
                List.of(
                    new ExcelRichTextRun("Quarterly", null),
                    new ExcelRichTextRun(
                        " Review",
                        new ExcelCellFont(
                            true, false, "Aptos", null, new ExcelColor("#112233"), null, null)))),
            new ExcelCommentAnchor(1, 2, 4, 6));
    var commentSnapshot =
        new ExcelCommentSnapshot(
            "Quarterly Review",
            "GridGrind",
            true,
            new ExcelRichTextSnapshot(
                List.of(
                    new ExcelRichTextRunSnapshot(
                        "Quarterly",
                        new ExcelCellFontSnapshot(
                            false,
                            false,
                            "Aptos",
                            ExcelFontHeight.fromPoints(java.math.BigDecimal.valueOf(11)),
                            null,
                            false,
                            false)),
                    new ExcelRichTextRunSnapshot(
                        " Review",
                        new ExcelCellFontSnapshot(
                            true,
                            false,
                            "Aptos",
                            ExcelFontHeight.fromPoints(java.math.BigDecimal.valueOf(11)),
                            new ExcelColorSnapshot("#112233"),
                            false,
                            false)))),
            new ExcelCommentAnchorSnapshot(1, 2, 4, 6));
    var formulaTarget = new ExcelNamedRangeTarget("SUM(Ops!$B$2:$B$5)");
    var table =
        new ExcelTableDefinition(
            "Queue",
            "Ops",
            "A1:B4",
            true,
            true,
            new ExcelTableStyle.None(),
            null,
            true,
            true,
            false,
            null,
            null,
            null,
            List.of(new ExcelTableColumnDefinition(1, null, null, " SUM ", "B2*2")));
    var protection = new ExcelWorkbookProtectionSettings(true, false, true, "book", "review");
    ExcelComment copiedComment = commentSnapshot.toAuthoringComment();

    assertEquals("SUM(Ops!$B$2:$B$5)", formulaTarget.refersToFormula());
    assertEquals(authoringComment.text(), copiedComment.text());
    assertEquals(authoringComment.author(), copiedComment.author());
    assertEquals(authoringComment.visible(), copiedComment.visible());
    assertEquals(authoringComment.anchor(), copiedComment.anchor());
    assertEquals(authoringComment.runs().plainText(), copiedComment.runs().plainText());
    assertEquals(new ExcelColor("#112233"), copiedComment.runs().runs().get(1).font().fontColor());
    assertEquals("", table.comment());
    assertEquals("", table.headerRowCellStyle());
    assertEquals("sum", table.columns().getFirst().totalsRowFunction());
    assertEquals("B2*2", table.columns().getFirst().calculatedColumnFormula());
    assertEquals("book", protection.workbookPassword());
    assertEquals("review", protection.revisionsPassword());
    assertEquals("A1", new ExcelNamedRangeTarget("Ops", "A1").range());
    assertEquals("Ops!$A$1", new ExcelNamedRangeTarget("Ops", "A1").refersToFormula());
    assertEquals("A1:B2", new ExcelNamedRangeTarget("Ops", "$A$1:$B$2").range());
    ExcelComment plainComment =
        new ExcelCommentSnapshot("Done", "GridGrind", false, null, null).toPlainComment();
    assertEquals("Done", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    ExcelComment plainAuthoringComment =
        new ExcelCommentSnapshot("Done", "GridGrind", false, null, null).toAuthoringComment();
    assertEquals("Done", plainAuthoringComment.text());
    assertNull(plainAuthoringComment.runs());
    assertNull(plainAuthoringComment.anchor());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelComment(
                "Quarterly Review",
                "GridGrind",
                true,
                new ExcelRichText(List.of(new ExcelRichTextRun("Mismatch", null))),
                null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelNamedRangeTarget(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelNamedRangeTarget("Ops", "A1", "SUM(Ops!$A$1)"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelWorkbookProtectionSettings(false, false, false, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelWorkbookProtectionSettings(false, false, false, null, " "));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelTableColumnDefinition(-1, "", "", "", ""));
    ExcelTableColumnDefinition normalizedColumn =
        new ExcelTableColumnDefinition(0, null, null, null, null);
    assertEquals("", normalizedColumn.uniqueName());
    assertEquals("", normalizedColumn.totalsRowLabel());
    assertEquals("", normalizedColumn.totalsRowFunction());
    assertEquals("", normalizedColumn.calculatedColumnFormula());
    ExcelTableDefinition normalizedTable =
        new ExcelTableDefinition(
            "Queue",
            "Ops",
            "A1:B2",
            false,
            true,
            new ExcelTableStyle.None(),
            null,
            false,
            false,
            false,
            null,
            null,
            null,
            null);
    assertEquals("", normalizedTable.comment());
    assertEquals(List.of(), normalizedTable.columns());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelTableDefinition(
                "Queue",
                "Ops",
                "A1:B2",
                false,
                true,
                new ExcelTableStyle.None(),
                "",
                false,
                false,
                false,
                "",
                "",
                "",
                Arrays.asList((ExcelTableColumnDefinition) null)));
  }

  @Test
  void conditionalFormattingAndSupportValueObjectsNormalizeAndValidate() throws Exception {
    for (ExcelConditionalFormattingIconSet iconSet : ExcelConditionalFormattingIconSet.values()) {
      assertEquals(iconSet, ExcelConditionalFormattingIconSet.fromPoi(iconSet.toPoi()));
      assertEquals(iconSet.toPoi().num, iconSet.thresholdCount());
    }
    for (ExcelConditionalFormattingThresholdType type :
        ExcelConditionalFormattingThresholdType.values()) {
      assertEquals(type, ExcelConditionalFormattingThresholdType.fromPoi(type.toPoi()));
    }

    var min =
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.MIN, null, null);
    var percentile =
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.PERCENTILE, null, 50.0d);
    var max =
        new ExcelConditionalFormattingThreshold(
            ExcelConditionalFormattingThresholdType.MAX, null, null);
    var colorScale =
        new ExcelConditionalFormattingRule.ColorScaleRule(
            List.of(min, percentile, max),
            List.of(
                new ExcelColor("#112233"), new ExcelColor("#445566"), new ExcelColor("#778899")),
            false);
    var dataBar =
        new ExcelConditionalFormattingRule.DataBarRule(
            new ExcelColor("#102030"), true, 10, 90, min, max, false);
    var iconSetRule =
        new ExcelConditionalFormattingRule.IconSetRule(
            ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
            true,
            true,
            List.of(
                new ExcelConditionalFormattingThreshold(
                    ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                percentile,
                max),
            false);
    var top10 = new ExcelConditionalFormattingRule.Top10Rule(10, true, false, false, null);
    var anchor = new ExcelCommentAnchor(1, 2, 3, 4);

    assertEquals(3, colorScale.thresholds().size());
    assertEquals(90, dataBar.widthMax());
    assertTrue(iconSetRule.iconOnly());
    assertEquals(10, top10.rank());
    assertEquals(1, anchor.firstColumn());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingThreshold(
                ExcelConditionalFormattingThresholdType.NUMBER, null, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                List.of(min), List.of(new ExcelColor("#112233")), false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                List.of(min, max), List.of(new ExcelColor("#112233")), false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.ColorScaleRule(
                null, List.of(new ExcelColor("#112233")), false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                new ExcelColor("#102030"), false, 10, 9, min, max, false));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelConditionalFormattingRule.DataBarRule(null, false, 0, 0, min, max, false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                new ExcelColor("#102030"), false, 0, 0, null, max, false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                new ExcelColor("#102030"), false, 0, 0, min, null, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                new ExcelColor("#102030"), false, -1, 0, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.DataBarRule(
                new ExcelColor("#102030"), false, 0, -1, min, max, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRule.IconSetRule(
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                false,
                List.of(min, max),
                false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelConditionalFormattingRule.Top10Rule(0, false, false, false, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchor(-1, 0, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchor(1, 0, 0, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchor(1, 2, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPrintMargins(-0.1d, 0, 0, 0, 0, 0));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertEquals("#AABBCC", ExcelRgbColorSupport.normalizeRgbHex("#aabbcc", "color"));
      assertThrows(
          NullPointerException.class, () -> ExcelRgbColorSupport.normalizeRgbHex("#AABBCC", null));
      assertThrows(
          IllegalArgumentException.class, () -> ExcelRgbColorSupport.normalizeRgbHex(" ", "color"));
      assertNull(ExcelRgbColorSupport.toRgbHex(null));
      assertEquals(
          "#112233",
          ExcelRgbColorSupport.toRgbHex(
              new XSSFColor(
                  new byte[] {0x11, 0x22, 0x33}, workbook.getStylesSource().getIndexedColors())));
      assertNull(
          ExcelRgbColorSupport.toRgbHex(
              new XSSFColor(
                  new byte[] {0x11, 0x22}, workbook.getStylesSource().getIndexedColors())));
      assertEquals(
          "#AABBCC",
          ExcelRgbColorSupport.toRgbHex(ExcelRgbColorSupport.toXssfColor(workbook, "#AABBCC")));
      assertThrows(
          NullPointerException.class, () -> ExcelRgbColorSupport.toXssfColor(null, "#AABBCC"));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelRgbColorSupport.toXssfColor(workbook, "#ABCDE"));

      workbook.createSheet("Source");
      workbook.createSheet("Replica");
      workbook.createSheet("Ledger");
      assertEquals(
          "",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook, "", org.apache.poi.ss.formula.FormulaType.CELL, 0, "Source", "Replica"));
      assertEquals(
          "SUM(Source!A1)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Source"));
      assertEquals(
          "SUM(Replica!A1,Replica:Ledger!B2)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1,Source:Ledger!B2)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica"));
      assertEquals(
          "SUM(Ledger:Replica!B2)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Ledger:Source!B2)",
              org.apache.poi.ss.formula.FormulaType.CELL,
              0,
              "Source",
              "Replica"));
      assertEquals(
          "SUM(Replica!A1)",
          ExcelFormulaSheetRenameSupport.renameSheet(
              workbook,
              "SUM(Source!A1)",
              org.apache.poi.ss.formula.FormulaType.CONDFORMAT,
              0,
              "Source",
              "Replica"));
    }
  }
}
