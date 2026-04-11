package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct tests for advanced mutation-oriented protocol DTO families. */
class AdvancedMutationProtocolTypesTest {
  @Test
  void workbookProtectionNamedRangeAndCommentInputsNormalizeAndValidate() {
    WorkbookProtectionInput protection =
        new WorkbookProtectionInput(null, true, null, "book-secret", "review-secret");

    assertFalse(protection.structureLocked());
    assertTrue(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertEquals("book-secret", protection.workbookPassword());
    assertEquals("review-secret", protection.revisionsPassword());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookProtectionInput(true, false, false, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookProtectionInput(true, false, false, null, " "));

    NamedRangeTarget explicit = new NamedRangeTarget("Budget", "C3:A1");
    NamedRangeTarget formula = new NamedRangeTarget("SUM(Budget!A1:A3)");
    assertEquals("Budget", explicit.sheetName());
    assertEquals("C3:A1", explicit.range());
    assertNull(explicit.formula());
    assertNull(formula.sheetName());
    assertNull(formula.range());
    assertEquals("SUM(Budget!A1:A3)", formula.formula());
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeTarget(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new NamedRangeTarget("Budget", "A1", "SUM(Budget!A1)"));

    CommentAnchorInput anchor = new CommentAnchorInput(1, 2, 4, 6);
    assertEquals(1, anchor.firstColumn());
    assertEquals(2, anchor.firstRow());
    assertEquals(4, anchor.lastColumn());
    assertEquals(6, anchor.lastRow());
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorInput(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorInput(1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorInput(0, 2, 1, 1));

    CommentInput richComment =
        new CommentInput(
            "Ada Lovelace",
            "GridGrind",
            true,
            List.of(new RichTextRunInput("Ada", null), new RichTextRunInput(" Lovelace", null)),
            anchor);
    assertTrue(richComment.visible());
    assertEquals(anchor, richComment.anchor());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CommentInput("Ada", "GridGrind", true, List.of(), null));
    assertThrows(
        NullPointerException.class,
        () -> new CommentInput("Ada", "GridGrind", true, List.of((RichTextRunInput) null), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CommentInput(
                "Ada Lovelace",
                "GridGrind",
                true,
                List.of(new RichTextRunInput("Ada", null)),
                null));
  }

  @Test
  void autofilterInputsNormalizeAndValidateAcrossAllCriterionFamilies() {
    AutofilterSortConditionInput colorSort =
        new AutofilterSortConditionInput("B2:B9", true, null, new ColorInput("#aabbcc"), null);
    AutofilterSortConditionInput iconSort =
        new AutofilterSortConditionInput("C2:C9", false, "ICON", null, 2);
    AutofilterSortStateInput sortState =
        new AutofilterSortStateInput("A1:F9", null, true, null, List.of(colorSort, iconSort));
    AutofilterSortStateInput sortStateWithMethod =
        new AutofilterSortStateInput("A1:F9", true, false, "PINYIN", List.of(colorSort));

    assertEquals("", colorSort.sortBy());
    assertEquals("#AABBCC", colorSort.color().rgb());
    assertEquals("ICON", iconSort.sortBy());
    assertEquals(2, iconSort.iconId());
    assertFalse(sortState.caseSensitive());
    assertTrue(sortState.columnSort());
    assertEquals("", sortState.sortMethod());
    assertEquals(List.of(colorSort, iconSort), sortState.conditions());
    assertEquals("PINYIN", sortStateWithMethod.sortMethod());
    assertEquals(
        new AutofilterFilterCriterionInput.Dynamic("TODAY", null, null),
        new AutofilterFilterCriterionInput.Dynamic("TODAY", null, null));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterSortConditionInput(null, false, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortStateInput(" ", false, false, null, List.of(colorSort)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionInput(" ", false, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionInput("A1:A2", false, null, null, -1));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterSortStateInput("A1:F9", false, false, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortStateInput("A1:F9", false, false, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterSortStateInput(
                "A1:F9", false, false, null, List.of((AutofilterSortConditionInput) null)));

    AutofilterFilterCriterionInput.Values values =
        new AutofilterFilterCriterionInput.Values(List.of("Queued", "Ready"), true);
    AutofilterFilterCriterionInput.Custom custom =
        new AutofilterFilterCriterionInput.Custom(
            true, List.of(new AutofilterFilterCriterionInput.CustomConditionInput("equal", "Ada")));
    AutofilterFilterCriterionInput.Dynamic dynamic =
        new AutofilterFilterCriterionInput.Dynamic("TODAY", 1.0d, 2.0d);
    AutofilterFilterCriterionInput.Top10 top10 =
        new AutofilterFilterCriterionInput.Top10(10, true, false);
    AutofilterFilterCriterionInput.Color color =
        new AutofilterFilterCriterionInput.Color(false, new ColorInput(null, 3, null, 0.25d));
    AutofilterFilterCriterionInput.Icon icon =
        new AutofilterFilterCriterionInput.Icon("3TrafficLights1", 2);
    AutofilterFilterColumnInput column = new AutofilterFilterColumnInput(2L, null, values);

    assertEquals(List.of("Queued", "Ready"), values.values());
    assertTrue(custom.and());
    assertEquals("TODAY", dynamic.type());
    assertEquals(10, top10.value());
    assertEquals(3, color.color().theme());
    assertEquals("3TrafficLights1", icon.iconSet());
    assertTrue(column.showButton());
    assertEquals(values, column.criterion());

    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionInput.Values(null, false));
    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterCriterionInput.Values(List.of("Queued", null), false));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionInput.Custom(true, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.CustomConditionInput(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.CustomConditionInput("equal", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Dynamic(" ", 1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Dynamic("TODAY", Double.POSITIVE_INFINITY, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Dynamic("TODAY", 1.0d, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Top10(0, true, false));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterCriterionInput.Color(true, null));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterCriterionInput.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionInput.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterColumnInput(-1L, true, values));
    assertThrows(
        NullPointerException.class, () -> new AutofilterFilterColumnInput(0L, false, null));
  }

  @Test
  void colorAndThresholdInputsNormalizeAndValidate() {
    ColorInput rgb = new ColorInput("#aabbcc");
    ColorInput themed = new ColorInput(null, 4, null, 0.5d);
    ColorInput indexed = new ColorInput(null, null, 64, null);
    assertEquals("#AABBCC", rgb.rgb());
    assertEquals(4, themed.theme());
    assertEquals(0.5d, themed.tint());
    assertEquals(64, indexed.indexed());
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(" ", null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput("123456", null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(null, -1, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ColorInput(null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColorInput("#AABBCC", null, null, Double.NEGATIVE_INFINITY));

    ConditionalFormattingThresholdInput threshold =
        new ConditionalFormattingThresholdInput(
            ExcelConditionalFormattingThresholdType.PERCENTILE, "A1*2", 90.0d);
    assertEquals(ExcelConditionalFormattingThresholdType.PERCENTILE, threshold.type());
    assertEquals("A1*2", threshold.formula());
    assertEquals(90.0d, threshold.value());
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingThresholdInput(null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingThresholdInput(
                ExcelConditionalFormattingThresholdType.NUMBER, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingThresholdInput(
                ExcelConditionalFormattingThresholdType.NUMBER, null, Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.ColorScaleRule(
                false,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MIN, null, null)),
                List.of(new ColorInput("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.ColorScaleRule(
                false,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                List.of(new ColorInput("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                new ColorInput("#112233"),
                false,
                -1,
                90,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                new ColorInput("#112233"),
                false,
                0,
                -1,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.DataBarRule(
                false,
                new ColorInput("#112233"),
                false,
                10,
                5,
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 0.0d),
                new ConditionalFormattingThresholdInput(
                    ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d)));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingRuleInput.IconSetRule(false, null, false, false, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.IconSetRule(
                false,
                ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
                false,
                false,
                List.of(
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 0.0d),
                    new ConditionalFormattingThresholdInput(
                        ExcelConditionalFormattingThresholdType.PERCENT, null, 50.0d))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleInput.Top10Rule(false, 0, true, false, null));
  }

  @Test
  void gradientAndStructuredStyleInputsNormalizeAndValidate() {
    CellGradientStopInput start = new CellGradientStopInput(0.0d, new ColorInput("#112233"));
    CellGradientStopInput finish =
        new CellGradientStopInput(1.0d, new ColorInput(null, 5, null, 0.2d));
    CellGradientFillInput gradient =
        new CellGradientFillInput(" path ", 45.0d, 0.1d, 0.2d, 0.3d, 0.4d, List.of(start, finish));
    assertEquals("PATH", gradient.type());
    assertEquals(List.of(start, finish), gradient.stops());
    assertEquals(
        "LINEAR",
        new CellGradientFillInput(null, null, null, null, null, null, List.of(start, finish))
            .type());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(-0.1d, new ColorInput("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientStopInput(Double.NaN, new ColorInput("#112233")));
    assertThrows(NullPointerException.class, () -> new CellGradientStopInput(0.5d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientFillInput(" ", null, null, null, null, null, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillInput(
                "diagonal", null, null, null, null, null, List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillInput(
                "LINEAR",
                Double.POSITIVE_INFINITY,
                null,
                null,
                null,
                null,
                List.of(start, finish)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientFillInput("LINEAR", null, null, null, null, null, List.of(start)));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellGradientFillInput(
                "LINEAR", null, null, null, null, null, List.of(start, null)));

    CellFontInput themedFont =
        new CellFontInput(null, null, null, null, null, 2, null, 0.4d, true, null);
    assertEquals(2, themedFont.fontColorTheme());
    assertEquals(0.4d, themedFont.fontColorTint());
    assertTrue(themedFont.underline());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, null, 0.4d, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, -1, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, -1, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontInput(
                null, null, null, null, null, null, null, Double.POSITIVE_INFINITY, null, null));

    CellFillInput gradientFill =
        new CellFillInput(null, null, null, null, null, null, null, null, null, gradient);
    CellFillInput themedFill =
        new CellFillInput(null, null, 2, null, 0.3d, null, null, null, null, null);
    assertEquals(gradient, gradientFill.gradient());
    assertEquals(2, themedFill.foregroundColorTheme());
    assertEquals(0.3d, themedFill.foregroundColorTint());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, 0.3d, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, -1, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, -1, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                null, null, null, null, Double.POSITIVE_INFINITY, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.SOLID, null, null, null, null, null, null, null, null, gradient));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.NONE, null, 2, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(null, null, null, null, null, null, 2, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFillInput(
                ExcelFillPattern.SOLID, "#112233", null, null, null, null, 2, null, null, null));

    CellBorderSideInput borderSide = new CellBorderSideInput(null, null, 1, null, 0.15d);
    assertEquals(1, borderSide.colorTheme());
    assertEquals(0.15d, borderSide.colorTint());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(null, null, null, null, 0.15d));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderSideInput(null, null, -1, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderSideInput(null, null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(null, null, null, null, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(ExcelBorderStyle.NONE, null, 1, null, null));
  }

  @Test
  void printAndTableInputsNormalizeAndValidate() {
    PrintMarginsInput margins = new PrintMarginsInput(0.5d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d);
    PrintSetupInput explicitSetup =
        new PrintSetupInput(margins, true, null, 9, true, null, 2, true, 4, List.of(6), List.of(3));
    PrintSetupInput defaultLikeSetup =
        new PrintSetupInput(null, null, null, null, null, null, null, null, null, null, null);

    assertEquals(margins, explicitSetup.margins());
    assertTrue(explicitSetup.horizontallyCentered());
    assertFalse(explicitSetup.verticallyCentered());
    assertEquals(9, explicitSetup.paperSize());
    assertTrue(explicitSetup.draft());
    assertFalse(explicitSetup.blackAndWhite());
    assertEquals(2, explicitSetup.copies());
    assertTrue(explicitSetup.useFirstPageNumber());
    assertEquals(4, explicitSetup.firstPageNumber());
    assertEquals(List.of(6), explicitSetup.rowBreaks());
    assertEquals(List.of(3), explicitSetup.columnBreaks());
    assertEquals(
        new PrintMarginsInput(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d), defaultLikeSetup.margins());
    assertEquals(0, defaultLikeSetup.paperSize());
    assertEquals(0, defaultLikeSetup.copies());
    assertEquals(0, defaultLikeSetup.firstPageNumber());
    assertEquals(PrintSetupInput.defaults().margins(), PrintSetupInput.defaults().margins());
    assertEquals(
        PrintSetupInput.defaults(),
        new PrintLayoutInput(null, null, null, null, null, null, null, null).setup());
    assertEquals(
        explicitSetup,
        new PrintLayoutInput(null, null, null, null, null, null, null, explicitSetup).setup());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsInput(-0.1d, 0.6d, 0.7d, 0.8d, 0.3d, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PrintMarginsInput(0.5d, 0.6d, 0.7d, 0.8d, Double.NaN, 0.2d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(margins, null, null, -1, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(margins, null, null, null, null, null, -1, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(margins, null, null, null, null, null, null, null, -1, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupInput(
                margins,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                List.of((Integer) null),
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupInput(
                margins, null, null, null, null, null, null, null, null, List.of(-1), null));

    TableColumnInput tableColumn = new TableColumnInput(0, null, null, " SUM ", null);
    TableColumnInput blankTotalsFunction = new TableColumnInput(0, null, null, " ", null);
    TableColumnInput nullTotalsFunction = new TableColumnInput(0, null, null, null, null);
    assertEquals("", tableColumn.uniqueName());
    assertEquals("", tableColumn.totalsRowLabel());
    assertEquals("sum", tableColumn.totalsRowFunction());
    assertEquals("", tableColumn.calculatedColumnFormula());
    assertEquals("", blankTotalsFunction.totalsRowFunction());
    assertEquals("", nullTotalsFunction.totalsRowFunction());
    assertThrows(
        IllegalArgumentException.class, () -> new TableColumnInput(-1, null, null, null, null));

    TableInput defaultedTable =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            null,
            null,
            new TableStyleInput.None(),
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
    TableInput explicitTable =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            false,
            false,
            new TableStyleInput.None(),
            "Table note",
            true,
            true,
            true,
            "HeaderStyle",
            "DataStyle",
            "TotalsStyle",
            List.of());
    assertFalse(defaultedTable.showTotalsRow());
    assertTrue(defaultedTable.hasAutofilter());
    assertEquals("", defaultedTable.comment());
    assertEquals(List.of(), defaultedTable.columns());
    assertEquals("Table note", explicitTable.comment());
    assertEquals("HeaderStyle", explicitTable.headerRowCellStyle());
    assertEquals("DataStyle", explicitTable.dataCellStyle());
    assertEquals("TotalsStyle", explicitTable.totalsRowCellStyle());
    assertThrows(
        NullPointerException.class,
        () ->
            new TableInput(
                "BudgetTable",
                "Budget",
                "A1:C4",
                false,
                true,
                new TableStyleInput.None(),
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                List.of((TableColumnInput) null)));
  }
}
