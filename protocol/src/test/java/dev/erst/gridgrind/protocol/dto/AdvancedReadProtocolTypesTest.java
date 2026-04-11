package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for the richer factual readback protocol types added for advanced XSSF parity. */
class AdvancedReadProtocolTypesTest {
  @Test
  void cellColorAndGradientReportsValidateRichReadShapes() {
    assertEquals(new CellColorReport("#ABCDEF"), new CellColorReport("#abcdef"));
    assertEquals(4, new CellColorReport(null, 4, null, 0.45d).theme());
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(" ", null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellColorReport("#12345G", null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(null, -1, null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellColorReport(null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellColorReport(null, null, null, Double.NaN));

    CellGradientStopReport firstStop = new CellGradientStopReport(0.0d, rgb("#112233"));
    CellGradientStopReport secondStop =
        new CellGradientStopReport(1.0d, new CellColorReport(null, 4, null, 0.45d));
    CellGradientFillReport gradient =
        new CellGradientFillReport(
            "LINEAR", 45.0d, 0.1d, 0.2d, 0.3d, 0.4d, List.of(firstStop, secondStop));

    assertEquals(2, gradient.stops().size());
    assertEquals(0.2d, gradient.right());
    assertEquals(
        new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), null),
        new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), null));
    assertEquals(
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellFillReport(ExcelFillPattern.NONE, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellGradientStopReport(1.5d, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillReport(
                " ", 45.0d, null, null, null, null, List.of(firstStop, secondStop)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellGradientFillReport(
                "LINEAR", Double.POSITIVE_INFINITY, null, null, null, null, List.of(firstStop)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGradientFillReport("LINEAR", 45.0d, null, null, null, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellGradientFillReport(
                "LINEAR", 45.0d, null, null, null, null, Arrays.asList(firstStop, null)));

    assertEquals(
        gradient, new CellFillReport(ExcelFillPattern.NONE, null, null, gradient).gradient());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillReport(ExcelFillPattern.NONE, rgb("#112233"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillReport(ExcelFillPattern.SOLID, rgb("#112233"), null, gradient));
  }

  @Test
  void commentPrintAndWorkbookProtectionReportsValidateRichReadShapes() {
    CellFontReport baseFont = font(rgb("#112233"));
    CommentAnchorReport anchor = new CommentAnchorReport(1, 2, 4, 6);
    GridGrindResponse.CommentReport plainComment =
        new GridGrindResponse.CommentReport("Plain", "Ada", false);
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport(
            "Hi there",
            "Ada",
            true,
            List.of(
                new RichTextRunReport("Hi ", baseFont),
                new RichTextRunReport("there", font(new CellColorReport(null, 5, null, 0.2d)))),
            anchor);

    assertNull(plainComment.runs());
    assertNull(plainComment.anchor());
    assertEquals(anchor, comment.anchor());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindResponse.CommentReport(
                "Mismatch", "Ada", true, List.of(new RichTextRunReport("Other", baseFont)), null));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(2, 0, 1, 0));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(0, 2, 0, 1));
    assertThrows(IllegalArgumentException.class, () -> new CommentAnchorReport(-1, 0, 1, 1));

    PrintSetupReport setup =
        new PrintSetupReport(
            new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
        () -> new PrintMarginsReport(-0.1d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
        NullPointerException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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
        IllegalArgumentException.class,
        () ->
            new PrintSetupReport(
                new PrintMarginsReport(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
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

    WorkbookProtectionReport protection =
        new WorkbookProtectionReport(true, false, true, true, false);
    assertTrue(protection.structureLocked());
  }

  @Test
  void autofilterReportsValidateAdvancedCriteriaAndSortShapes() {
    AutofilterFilterCriterionReport.Values values =
        new AutofilterFilterCriterionReport.Values(List.of("Queued", "Blocked"), true);
    AutofilterFilterCriterionReport.Custom custom =
        new AutofilterFilterCriterionReport.Custom(
            true,
            List.of(new AutofilterFilterCriterionReport.CustomConditionReport("equal", "Ada")));
    AutofilterFilterCriterionReport.Dynamic dynamic =
        new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, 2.0d);
    AutofilterFilterCriterionReport.Top10 top10 =
        new AutofilterFilterCriterionReport.Top10(true, false, 10.0d, 8.0d);
    AutofilterFilterCriterionReport.Color color =
        new AutofilterFilterCriterionReport.Color(false, new CellColorReport(null, 4, null, 0.45d));
    AutofilterFilterCriterionReport.Icon icon =
        new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2);
    AutofilterSortConditionReport sortCondition =
        new AutofilterSortConditionReport("A2:A5", true, null, rgb("#AABBCC"), 1);
    AutofilterSortStateReport sortState =
        new AutofilterSortStateReport("A1:F5", true, false, null, List.of(sortCondition));
    AutofilterEntryReport.SheetOwned sheetOwned =
        new AutofilterEntryReport.SheetOwned(
            "A1:F5",
            List.of(
                new AutofilterFilterColumnReport(0L, false, values),
                new AutofilterFilterColumnReport(1L, true, custom),
                new AutofilterFilterColumnReport(2L, true, dynamic),
                new AutofilterFilterColumnReport(3L, true, top10),
                new AutofilterFilterColumnReport(4L, true, color),
                new AutofilterFilterColumnReport(5L, true, icon)),
            sortState);
    AutofilterEntryReport.TableOwned tableOwned =
        new AutofilterEntryReport.TableOwned("H1:I5", "QueueTable", List.of(), sortState);
    AutofilterEntryReport.TableOwned tableOwnedWithColumn =
        new AutofilterEntryReport.TableOwned(
            "N1:O5",
            "QueueMirror",
            List.of(new AutofilterFilterColumnReport(9L, true, values)),
            sortState);
    AutofilterEntryReport.SheetOwned defaultSheetOwned =
        new AutofilterEntryReport.SheetOwned("J1:K4");
    AutofilterEntryReport.TableOwned defaultTableOwned =
        new AutofilterEntryReport.TableOwned("L1:M4", "AuditTable");
    AutofilterFilterCriterionReport.Values emptyValues =
        new AutofilterFilterCriterionReport.Values(List.of(), false);

    assertEquals(sortState, sheetOwned.sortState());
    assertEquals("QueueTable", tableOwned.tableName());
    assertEquals(1, tableOwnedWithColumn.filterColumns().size());
    assertEquals(List.of(), emptyValues.values());
    assertEquals(List.of(), defaultSheetOwned.filterColumns());
    assertNull(defaultSheetOwned.sortState());
    assertEquals(List.of(), defaultTableOwned.filterColumns());
    assertNull(defaultTableOwned.sortState());
    assertEquals("", sortCondition.sortBy());
    assertEquals("", sortState.sortMethod());

    assertThrows(
        NullPointerException.class,
        () -> new AutofilterFilterCriterionReport.Values(Arrays.asList("Queued", null), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Custom(true, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.CustomConditionReport(" ", "Ada"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic(" ", 1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic("TODAY", Double.POSITIVE_INFINITY, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Dynamic("TODAY", 1.0d, Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Top10(true, false, -1.0d, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new AutofilterFilterCriterionReport.Top10(
                true, false, 10.0d, Double.NEGATIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterCriterionReport.Icon(" ", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.Icon("3TrafficLights1", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new AutofilterFilterColumnReport(-1L, true, values));
    assertEquals(
        "CELL_COLOR",
        new AutofilterSortConditionReport("A2:A5", false, "CELL_COLOR", null, null).sortBy());
    assertEquals(" ", new AutofilterSortConditionReport(" ", false, null, null, null).range());
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterSortConditionReport("A2:A5", true, null, rgb("#AABBCC"), -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterFilterCriterionReport.CustomConditionReport("equal", " "));
    assertEquals(
        " ", new AutofilterSortStateReport(" ", true, false, null, List.of(sortCondition)).range());
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterSortStateReport(
                "A1:F5", true, false, null, Arrays.asList(sortCondition, null)));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterEntryReport.SheetOwned(
                "A1:F5", Arrays.asList(sheetOwned.filterColumns().getFirst(), null), sortState));
    assertThrows(
        NullPointerException.class,
        () ->
            new AutofilterEntryReport.TableOwned(
                "A1:F5",
                "QueueTable",
                Arrays.asList(sheetOwned.filterColumns().getFirst(), null),
                sortState));
  }

  @Test
  void tableConditionalFormattingAndWorkbookProtectionReadContractsValidate() {
    TableEntryReport normalized =
        new TableEntryReport(
            "BudgetTable",
            "Budget",
            "A1:B5",
            1,
            1,
            List.of("Item", "Amount"),
            List.of(
                new TableColumnReport(1L, "Item", null, null, null, null),
                new TableColumnReport(2L, "Amount", "UniqueAmount", "Total", "sum", "[@Amount]*2")),
            new TableStyleReport.Named("TableStyleMedium2", false, false, true, false),
            true,
            null,
            true,
            true,
            false,
            null,
            null,
            null);

    assertEquals("", normalized.comment());
    assertEquals("", normalized.headerRowCellStyle());
    assertEquals("UniqueAmount", normalized.columns().get(1).uniqueName());
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableColumnReport(-1L, "Item", null, null, null, null));

    DifferentialStyleReport style =
        new DifferentialStyleReport(
            "0.00", true, null, null, "#AABBCC", null, null, null, null, List.of());
    ConditionalFormattingRuleReport.Top10Rule top10 =
        new ConditionalFormattingRuleReport.Top10Rule(1, false, 10, true, false, style);

    assertEquals(10, top10.rank());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleReport.Top10Rule(1, false, -1, false, false, style));

    WorkbookReadOperation.GetWorkbookProtection read =
        new WorkbookReadOperation.GetWorkbookProtection("workbook-protection");
    WorkbookReadResult.WorkbookProtectionResult result =
        new WorkbookReadResult.WorkbookProtectionResult(
            "workbook-protection", new WorkbookProtectionReport(true, false, true, true, false));

    assertEquals("workbook-protection", read.requestId());
    assertTrue(result.protection().workbookPasswordHashPresent());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadOperation.GetWorkbookProtection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadOperation.GetWorkbookProtection(" "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.WorkbookProtectionResult("workbook-protection", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookProtectionResult(
                " ", new WorkbookProtectionReport(true, false, true, true, false)));
  }

  private static CellColorReport rgb(String rgb) {
    return new CellColorReport(rgb);
  }

  private static CellFontReport font(CellColorReport color) {
    return new CellFontReport(
        false,
        false,
        "Aptos",
        new FontHeightReport(220, BigDecimal.valueOf(11)),
        color,
        false,
        false);
  }
}
