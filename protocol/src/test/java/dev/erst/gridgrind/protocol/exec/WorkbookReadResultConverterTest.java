package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for converting advanced engine read results into protocol response shapes. */
class WorkbookReadResultConverterTest {
  @Test
  void convertsPlainCommentAndWorkbookProtectionFactsDirectly() {
    assertNull(WorkbookReadResultConverter.toCommentReport((ExcelComment) null));
    GridGrindResponse.CommentReport plainComment =
        WorkbookReadResultConverter.toCommentReport(new ExcelComment("Review", "GridGrind", false));
    WorkbookProtectionReport protection =
        WorkbookReadResultConverter.toWorkbookProtectionReport(
            new ExcelWorkbookProtectionSnapshot(true, false, true, true, false));

    assertEquals("Review", plainComment.text());
    assertEquals("GridGrind", plainComment.author());
    assertNull(plainComment.runs());
    assertTrue(protection.structureLocked());
    assertTrue(protection.revisionLocked());
  }

  @Test
  void convertsAdvancedReadResultsIntoProtocolShapes() {
    WorkbookReadResult.WorkbookProtectionResult protection =
        assertInstanceOf(
            WorkbookReadResult.WorkbookProtectionResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookProtectionResult(
                    "workbook-protection",
                    new ExcelWorkbookProtectionSnapshot(true, false, true, true, false))));
    assertTrue(protection.protection().structureLocked());

    WorkbookReadResult.CellsResult cells =
        assertInstanceOf(
            WorkbookReadResult.CellsResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                    "cells", "Budget", List.of(advancedCell()))));
    GridGrindResponse.CellReport.TextReport cell =
        assertInstanceOf(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst());
    assertEquals(new CellColorReport(null, 2, null, 0.25d), cell.style().font().fontColor());
    assertEquals(new CellColorReport(null, null, 12, null), cell.style().border().bottom().color());
    assertNotNull(cell.style().fill().gradient());
    assertEquals(2, cell.style().fill().gradient().stops().size());
    assertEquals("https://example.com/tasks/42", ((HyperlinkTarget.Url) cell.hyperlink()).target());
    assertNotNull(cell.comment());
    assertEquals(2, cell.comment().runs().size());
    assertEquals(1, cell.comment().anchor().firstColumn());

    WorkbookReadResult.CommentsResult comments =
        assertInstanceOf(
            WorkbookReadResult.CommentsResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.CellComment(
                            "C3", richComment())))));
    assertEquals(2, comments.comments().getFirst().comment().runs().size());
    assertEquals(6, comments.comments().getFirst().comment().anchor().lastRow());

    WorkbookReadResult.PrintLayoutResult printLayout =
        assertInstanceOf(
            WorkbookReadResult.PrintLayoutResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                    "print-layout", "Budget", advancedPrintLayout())));
    assertEquals(List.of(10, 20), printLayout.layout().setup().rowBreaks());
    assertEquals(9, printLayout.layout().setup().paperSize());

    WorkbookReadResult.AutofiltersResult autofilters =
        assertInstanceOf(
            WorkbookReadResult.AutofiltersResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult(
                    "autofilters", "Budget", advancedAutofilters())));
    AutofilterEntryReport.SheetOwned sheetOwned =
        assertInstanceOf(
            AutofilterEntryReport.SheetOwned.class, autofilters.autofilters().getFirst());
    assertEquals(6, sheetOwned.filterColumns().size());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Custom.class,
        sheetOwned.filterColumns().get(1).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Dynamic.class,
        sheetOwned.filterColumns().get(2).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Top10.class, sheetOwned.filterColumns().get(3).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Color.class, sheetOwned.filterColumns().get(4).criterion());
    assertInstanceOf(
        AutofilterFilterCriterionReport.Icon.class, sheetOwned.filterColumns().get(5).criterion());
    assertEquals("A1:F5", sheetOwned.sortState().range());
    assertInstanceOf(AutofilterEntryReport.TableOwned.class, autofilters.autofilters().get(1));

    WorkbookReadResult.TablesResult tables =
        assertInstanceOf(
            WorkbookReadResult.TablesResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult(
                    "tables", List.of(advancedTable()))));
    assertEquals("HeaderStyle", tables.tables().getFirst().headerRowCellStyle());
    assertEquals("Total", tables.tables().getFirst().columns().get(1).totalsRowLabel());

    WorkbookReadResult.ConditionalFormattingResult conditionalFormatting =
        assertInstanceOf(
            WorkbookReadResult.ConditionalFormattingResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult(
                    "conditional-formatting",
                    "Budget",
                    List.of(
                        new ExcelConditionalFormattingBlockSnapshot(
                            List.of("A1:A5"),
                            List.of(
                                new ExcelConditionalFormattingRuleSnapshot.Top10Rule(
                                    1, false, 10, true, false, differentialStyle())))))));
    assertInstanceOf(
        ConditionalFormattingRuleReport.Top10Rule.class,
        conditionalFormatting.conditionalFormattingBlocks().getFirst().rules().getFirst());
  }

  private static ExcelCellSnapshot advancedCell() {
    ExcelRichTextSnapshot richText =
        new ExcelRichTextSnapshot(List.of(new ExcelRichTextRunSnapshot("Styled", advancedFont())));
    return new ExcelCellSnapshot.TextSnapshot(
        "C3",
        "STRING",
        "Styled",
        advancedStyle(),
        ExcelCellMetadataSnapshot.of(
            new ExcelHyperlink.Url("https://example.com/tasks/42"), richComment()),
        "Styled",
        richText);
  }

  private static ExcelCommentSnapshot richComment() {
    return new ExcelCommentSnapshot(
        "Hi there",
        "Ada",
        true,
        new ExcelRichTextSnapshot(
            List.of(
                new ExcelRichTextRunSnapshot("Hi ", advancedFont()),
                new ExcelRichTextRunSnapshot("there", accentFont()))),
        new ExcelCommentAnchorSnapshot(1, 2, 4, 6));
  }

  private static ExcelCellStyleSnapshot advancedStyle() {
    return new ExcelCellStyleSnapshot(
        "0.00",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER, 0, 0),
        advancedFont(),
        new ExcelCellFillSnapshot(
            ExcelFillPattern.NONE,
            null,
            null,
            new ExcelGradientFillSnapshot(
                "LINEAR",
                45.0d,
                null,
                null,
                null,
                null,
                List.of(
                    new ExcelGradientStopSnapshot(0.0d, new ExcelColorSnapshot("#112233")),
                    new ExcelGradientStopSnapshot(
                        1.0d, new ExcelColorSnapshot(null, 4, null, 0.45d))))),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(
                ExcelBorderStyle.THICK, new ExcelColorSnapshot(null, null, 12, null)),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private static ExcelCellFontSnapshot advancedFont() {
    return new ExcelCellFontSnapshot(
        true,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        new ExcelColorSnapshot(null, 2, null, 0.25d),
        false,
        false);
  }

  private static ExcelCellFontSnapshot accentFont() {
    return new ExcelCellFontSnapshot(
        false,
        false,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        new ExcelColorSnapshot("#AABBCC"),
        false,
        false);
  }

  private static ExcelPrintLayoutSnapshot advancedPrintLayout() {
    return new ExcelPrintLayoutSnapshot(
        new ExcelPrintLayout(
            new ExcelPrintLayout.Area.Range("A1:F20"),
            ExcelPrintOrientation.LANDSCAPE,
            new ExcelPrintLayout.Scaling.Fit(1, 0),
            new ExcelPrintLayout.TitleRows.Band(0, 0),
            new ExcelPrintLayout.TitleColumns.None(),
            new ExcelHeaderFooterText("Ops", "Queue", ""),
            new ExcelHeaderFooterText("", "Internal", "Page &P")),
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(0.5d, 0.5d, 1.0d, 1.0d, 0.3d, 0.3d),
            true,
            false,
            9,
            false,
            true,
            2,
            true,
            3,
            List.of(10, 20),
            List.of(2, 4)));
  }

  private static List<ExcelAutofilterSnapshot> advancedAutofilters() {
    ExcelAutofilterSortStateSnapshot sortState =
        new ExcelAutofilterSortStateSnapshot(
            "A1:F5",
            true,
            false,
            "",
            List.of(
                new ExcelAutofilterSortConditionSnapshot(
                    "A2:A5", true, "", new ExcelColorSnapshot("#AABBCC"), 1)));
    List<ExcelAutofilterFilterColumnSnapshot> filterColumns =
        List.of(
            new ExcelAutofilterFilterColumnSnapshot(
                0L,
                false,
                new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Queued"), true)),
            new ExcelAutofilterFilterColumnSnapshot(
                1L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Custom(
                    true,
                    List.of(
                        new ExcelAutofilterFilterCriterionSnapshot.CustomCondition(
                            "equal", "Ada")))),
            new ExcelAutofilterFilterColumnSnapshot(
                2L, true, new ExcelAutofilterFilterCriterionSnapshot.Dynamic("TODAY", 1.0d, 2.0d)),
            new ExcelAutofilterFilterColumnSnapshot(
                3L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 10.0d, 8.0d)),
            new ExcelAutofilterFilterColumnSnapshot(
                4L,
                true,
                new ExcelAutofilterFilterCriterionSnapshot.Color(
                    false, new ExcelColorSnapshot(null, 4, null, 0.45d))),
            new ExcelAutofilterFilterColumnSnapshot(
                5L, true, new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2)));
    return List.of(
        new ExcelAutofilterSnapshot.SheetOwned("A1:F5", filterColumns, sortState),
        new ExcelAutofilterSnapshot.TableOwned("H1:I5", "QueueTable", List.of(), sortState));
  }

  private static ExcelTableSnapshot advancedTable() {
    return new ExcelTableSnapshot(
        "QueueTable",
        "Budget",
        "A1:B5",
        1,
        1,
        List.of("Item", "Amount"),
        List.of(
            new ExcelTableColumnSnapshot(1L, "Item", "", "", "", ""),
            new ExcelTableColumnSnapshot(
                2L, "Amount", "UniqueAmount", "Total", "sum", "[@Amount]*2")),
        new ExcelTableStyleSnapshot.Named("TableStyleMedium2", false, false, true, false),
        true,
        "Queue comment",
        true,
        true,
        false,
        "HeaderStyle",
        "DataStyle",
        "TotalsStyle");
  }

  private static ExcelDifferentialStyleSnapshot differentialStyle() {
    return new ExcelDifferentialStyleSnapshot(
        "0.00", true, null, null, "#AABBCC", null, null, null, null, List.of());
  }
}
