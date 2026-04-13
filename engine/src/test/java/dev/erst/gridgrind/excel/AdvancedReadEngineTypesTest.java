package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Tests for richer factual engine snapshot types added for advanced XSSF parity reads. */
class AdvancedReadEngineTypesTest {
  @Test
  void colorGradientAndFillSnapshotsValidateRichReadShapes() {
    assertEquals(new ExcelColorSnapshot("#ABCDEF"), new ExcelColorSnapshot("#abcdef"));
    assertEquals(4, new ExcelColorSnapshot(null, 4, null, 0.45d).theme());
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelColorSnapshot(null, null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelColorSnapshot(" ", null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelColorSnapshot("#12345G", null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelColorSnapshot(null, -1, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelColorSnapshot(null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelColorSnapshot(null, null, null, Double.NaN));

    ExcelGradientStopSnapshot firstStop = new ExcelGradientStopSnapshot(0.0d, rgb("#112233"));
    ExcelGradientStopSnapshot secondStop =
        new ExcelGradientStopSnapshot(1.0d, new ExcelColorSnapshot(null, 4, null, 0.45d));
    ExcelGradientFillSnapshot gradient =
        new ExcelGradientFillSnapshot(
            "LINEAR", 45.0d, 0.1d, 0.2d, 0.3d, 0.4d, List.of(firstStop, secondStop));

    assertEquals(2, gradient.stops().size());
    assertEquals(0.2d, gradient.right());
    assertEquals(
        new ExcelCellFillSnapshot(ExcelFillPattern.SOLID, rgb("#112233"), null),
        new ExcelCellFillSnapshot(ExcelFillPattern.SOLID, rgb("#112233"), null));
    assertEquals(
        new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, null),
        new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelGradientStopSnapshot(1.5d, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelGradientFillSnapshot(
                " ", 45.0d, null, null, null, null, List.of(firstStop, secondStop)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelGradientFillSnapshot(
                "LINEAR", Double.POSITIVE_INFINITY, null, null, null, null, List.of(firstStop)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelGradientFillSnapshot("LINEAR", 45.0d, null, null, null, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelGradientFillSnapshot(
                "LINEAR", 45.0d, null, null, null, null, Arrays.asList(firstStop, null)));

    assertEquals(
        gradient,
        new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, null, gradient).gradient());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFillSnapshot(ExcelFillPattern.NONE, rgb("#112233"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFillSnapshot(ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFillSnapshot(ExcelFillPattern.SOLID, rgb("#112233"), null, gradient));
  }

  @Test
  void colorSnapshotSupportAndRgbHelpersPreserveWorkbookColorSemantics() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFColor rgbColor =
          new XSSFColor(
              new byte[] {0x11, 0x22, 0x33}, workbook.getStylesSource().getIndexedColors());
      XSSFColor indexedColor = new XSSFColor();
      indexedColor.setIndexed(IndexedColors.DARK_RED.getIndex());
      XSSFColor tintedTheme = new XSSFColor();
      tintedTheme.setTheme(3);
      tintedTheme.setTint(0.25d);
      XSSFColor malformedRgb =
          new XSSFColor(new byte[] {0x11, 0x22}, workbook.getStylesSource().getIndexedColors());
      CTColor themed = CTColor.Factory.newInstance();
      themed.setTheme(2L);
      CTColor indexed = CTColor.Factory.newInstance();
      indexed.setIndexed(7L);
      CTColor argb = CTColor.Factory.newInstance();
      argb.setRgb(new byte[] {(byte) 0xFF, 0x44, 0x55, 0x66});

      assertNull(ExcelColorSnapshotSupport.snapshot((XSSFColor) null));
      assertNull(ExcelColorSnapshotSupport.snapshot(new XSSFColor()));
      assertEquals(new ExcelColorSnapshot("#112233"), ExcelColorSnapshotSupport.snapshot(rgbColor));
      assertEquals(
          new ExcelColorSnapshot(null, null, (int) IndexedColors.DARK_RED.getIndex(), null),
          ExcelColorSnapshotSupport.snapshot(indexedColor));
      assertEquals(
          new ExcelColorSnapshot(null, 3, null, 0.25d),
          ExcelColorSnapshotSupport.snapshot(tintedTheme));
      assertEquals("#112233", ExcelRgbColorSupport.toRgbHex(rgbColor));
      assertNull(ExcelRgbColorSupport.toRgbHex(null));
      assertNull(ExcelRgbColorSupport.toRgbHex(malformedRgb));
      assertNull(ExcelColorSnapshotSupport.snapshot(workbook, null));
      assertEquals(
          new ExcelColorSnapshot(null, 2, null, null),
          ExcelColorSnapshotSupport.snapshot(workbook, themed));
      assertEquals(
          new ExcelColorSnapshot(null, null, 7, null),
          ExcelColorSnapshotSupport.snapshot(workbook, indexed));
      assertEquals(
          new ExcelColorSnapshot("#445566"), ExcelColorSnapshotSupport.snapshot(workbook, argb));
      assertNull(ExcelColorSnapshotSupport.snapshot(workbook, CTColor.Factory.newInstance()));
    }
  }

  @Test
  void commentPrintAndProtectionSnapshotsValidateRichReadShapes() {
    ExcelCellFontSnapshot baseFont = font(rgb("#112233"));
    ExcelCommentAnchorSnapshot anchor = new ExcelCommentAnchorSnapshot(1, 2, 4, 6);
    ExcelRichTextSnapshot runs =
        new ExcelRichTextSnapshot(
            List.of(
                new ExcelRichTextRunSnapshot("Hi ", baseFont),
                new ExcelRichTextRunSnapshot(
                    "there", font(new ExcelColorSnapshot(null, 5, null, 0.2d)))));
    ExcelCommentSnapshot comment = new ExcelCommentSnapshot("Hi there", "Ada", true, runs, anchor);

    assertEquals(anchor, comment.anchor());
    assertEquals(new ExcelComment("Hi there", "Ada", true), comment.toPlainComment());
    assertEquals(
        new ExcelComment("Ok", "Ada", false),
        new ExcelCommentSnapshot("Ok", "Ada", false, null, null).toPlainComment());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCommentSnapshot(
                "Mismatch",
                "Ada",
                true,
                new ExcelRichTextSnapshot(List.of(new ExcelRichTextRunSnapshot("Other", baseFont))),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCommentSnapshot(" ", "Ada", true, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCommentSnapshot("Hi there", " ", true, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchorSnapshot(2, 0, 1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchorSnapshot(0, 2, 0, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCommentAnchorSnapshot(-1, 0, 1, 1));

    ExcelPrintSetupSnapshot setup =
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
            false,
            true,
            false,
            9,
            false,
            true,
            2,
            true,
            3,
            List.of(10, 20),
            List.of(2, 4));

    assertEquals(List.of(10, 20), setup.rowBreaks());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPrintMarginsSnapshot(-0.1d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                -1,
                false,
                true,
                2,
                true,
                3,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                -1,
                true,
                3,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                -1,
                List.of(10),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                List.of(-1),
                List.of(2)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                List.of(10),
                List.of(-1)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
                false,
                true,
                false,
                9,
                false,
                true,
                2,
                true,
                3,
                Arrays.asList(10, null),
                List.of(2)));

    ExcelWorkbookProtectionSnapshot protection =
        new ExcelWorkbookProtectionSnapshot(true, false, true, true, false);
    WorkbookReadCommand.GetWorkbookProtection read =
        new WorkbookReadCommand.GetWorkbookProtection("workbook-protection");
    WorkbookReadResult.WorkbookProtectionResult result =
        new WorkbookReadResult.WorkbookProtectionResult("workbook-protection", protection);

    assertTrue(protection.structureLocked());
    assertEquals("workbook-protection", read.requestId());
    assertTrue(result.protection().workbookPasswordHashPresent());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadCommand.GetWorkbookProtection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadCommand.GetWorkbookProtection(" "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.WorkbookProtectionResult("workbook-protection", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookProtectionResult(
                " ", new ExcelWorkbookProtectionSnapshot(true, false, true, true, false)));
  }

  @Test
  void autofilterAndTableSnapshotsValidateAdvancedCriteriaAndSortShapes() {
    ExcelAutofilterFilterCriterionSnapshot.Values values =
        new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Queued", "Blocked"), true);
    ExcelAutofilterFilterCriterionSnapshot.Custom custom =
        new ExcelAutofilterFilterCriterionSnapshot.Custom(
            true,
            List.of(new ExcelAutofilterFilterCriterionSnapshot.CustomCondition("equal", "Ada")));
    ExcelAutofilterFilterCriterionSnapshot.Dynamic dynamic =
        new ExcelAutofilterFilterCriterionSnapshot.Dynamic("TODAY", 1.0d, 2.0d);
    ExcelAutofilterFilterCriterionSnapshot.Top10 top10 =
        new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 10.0d, 8.0d);
    ExcelAutofilterFilterCriterionSnapshot.Color color =
        new ExcelAutofilterFilterCriterionSnapshot.Color(
            false, new ExcelColorSnapshot(null, 4, null, 0.45d));
    ExcelAutofilterFilterCriterionSnapshot.Icon icon =
        new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2);
    ExcelAutofilterSortConditionSnapshot sortCondition =
        new ExcelAutofilterSortConditionSnapshot("A2:A5", true, null, rgb("#AABBCC"), 1);
    ExcelAutofilterSortStateSnapshot sortState =
        new ExcelAutofilterSortStateSnapshot("A1:F5", true, false, null, List.of(sortCondition));
    ExcelAutofilterSnapshot.SheetOwned sheetOwned =
        new ExcelAutofilterSnapshot.SheetOwned(
            "A1:F5",
            List.of(
                new ExcelAutofilterFilterColumnSnapshot(0L, false, values),
                new ExcelAutofilterFilterColumnSnapshot(1L, true, custom),
                new ExcelAutofilterFilterColumnSnapshot(2L, true, dynamic),
                new ExcelAutofilterFilterColumnSnapshot(3L, true, top10),
                new ExcelAutofilterFilterColumnSnapshot(4L, true, color),
                new ExcelAutofilterFilterColumnSnapshot(5L, true, icon)),
            sortState);
    ExcelAutofilterSnapshot.TableOwned tableOwned =
        new ExcelAutofilterSnapshot.TableOwned("H1:I5", "QueueTable", List.of(), sortState);
    ExcelAutofilterSnapshot.TableOwned tableOwnedWithColumn =
        new ExcelAutofilterSnapshot.TableOwned(
            "N1:O5",
            "QueueMirror",
            List.of(new ExcelAutofilterFilterColumnSnapshot(9L, true, values)),
            sortState);
    ExcelAutofilterSnapshot.TableOwned defaultTableOwned =
        new ExcelAutofilterSnapshot.TableOwned("L1:M4", "AuditTable");
    ExcelAutofilterFilterCriterionSnapshot.Values emptyValues =
        new ExcelAutofilterFilterCriterionSnapshot.Values(List.of(), false);

    assertEquals(sortState, sheetOwned.sortState());
    assertEquals("QueueTable", tableOwned.tableName());
    assertEquals(1, tableOwnedWithColumn.filterColumns().size());
    assertEquals(List.of(), emptyValues.values());
    assertEquals(List.of(), defaultTableOwned.filterColumns());
    assertNull(defaultTableOwned.sortState());
    assertEquals("", sortCondition.sortBy());
    assertEquals("", sortState.sortMethod());

    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelAutofilterFilterCriterionSnapshot.Values(
                Arrays.asList("Queued", null), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.CustomCondition("equal", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.Dynamic(" ", 1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelAutofilterFilterCriterionSnapshot.Dynamic(
                "TODAY", Double.POSITIVE_INFINITY, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.Dynamic("TODAY", 1.0d, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, -1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelAutofilterFilterCriterionSnapshot.Top10(
                true, false, 10.0d, Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterFilterColumnSnapshot(-1L, true, values));
    assertEquals(
        "CELL_COLOR",
        new ExcelAutofilterSortConditionSnapshot("A2:A5", false, "CELL_COLOR", null, null)
            .sortBy());
    assertEquals(
        " ", new ExcelAutofilterSortConditionSnapshot(" ", false, null, null, null).range());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelAutofilterSortConditionSnapshot("A2:A5", true, null, rgb("#AABBCC"), -1));
    assertEquals(
        " ",
        new ExcelAutofilterSortStateSnapshot(" ", true, false, null, List.of(sortCondition))
            .range());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelAutofilterSortStateSnapshot(
                "A1:F5", true, false, null, Arrays.asList(sortCondition, null)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelAutofilterSnapshot.TableOwned(
                "A1:F5",
                "QueueTable",
                Arrays.asList(sheetOwned.filterColumns().getFirst(), null),
                sortState));
  }

  @Test
  void tableAndConditionalFormattingSnapshotsValidateAdvancedMetadata() {
    ExcelTableSnapshot normalized =
        new ExcelTableSnapshot(
            "BudgetTable",
            "Budget",
            "A1:B5",
            1,
            1,
            List.of("Item", "Amount"),
            List.of(
                new ExcelTableColumnSnapshot(1L, "Item", null, null, null, null),
                new ExcelTableColumnSnapshot(
                    2L, "Amount", "UniqueAmount", "Total", "sum", "[@Amount]*2")),
            new ExcelTableStyleSnapshot.Named("TableStyleMedium2", false, false, true, false),
            true,
            null,
            true,
            true,
            false,
            null,
            null,
            null);
    ExcelConditionalFormattingRuleSnapshot.Top10Rule top10 =
        new ExcelConditionalFormattingRuleSnapshot.Top10Rule(
            1,
            false,
            10,
            true,
            false,
            new ExcelDifferentialStyleSnapshot(
                "0.00", true, null, null, "#AABBCC", null, null, null, null, List.of()));

    assertEquals("", normalized.comment());
    assertEquals("", normalized.headerRowCellStyle());
    assertEquals("UniqueAmount", normalized.columns().get(1).uniqueName());
    assertEquals(10, top10.rank());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelTableColumnSnapshot(-1L, "Item", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.Top10Rule(1, false, -1, false, false, null));
  }

  private static ExcelColorSnapshot rgb(String rgb) {
    return new ExcelColorSnapshot(rgb);
  }

  private static ExcelCellFontSnapshot font(ExcelColorSnapshot color) {
    return new ExcelCellFontSnapshot(
        false,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
        color,
        false,
        false);
  }
}
