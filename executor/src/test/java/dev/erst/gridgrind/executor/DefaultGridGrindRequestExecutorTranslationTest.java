package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Command, read-model, and diagnostic translation tests for DefaultGridGrindRequestExecutor. */
class DefaultGridGrindRequestExecutorTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void extractsSheetOnlyContextForDeleteSheetOperations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation ensureSheet =
        mutate(new SheetSelector.ByName("Archive"), new MutationAction.EnsureSheet());
    ExecutorTestPlanSupport.PendingMutation deleteSheet =
        mutate(new SheetSelector.ByName("Archive"), new MutationAction.DeleteSheet());

    assertNull(formulaFor(ensureSheet, exception));
    assertNull(formulaFor(deleteSheet, exception));
    assertEquals("Archive", sheetNameFor(ensureSheet, exception));
    assertEquals("Archive", sheetNameFor(deleteSheet, exception));
    assertNull(addressFor(ensureSheet, exception));
    assertNull(addressFor(deleteSheet, exception));
    assertNull(rangeFor(ensureSheet, exception));
    assertNull(rangeFor(deleteSheet, exception));
  }

  @Test
  void extractsSheetStateContextForB1Operations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation copySheet =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.CopySheet("Budget Copy", new SheetCopyPosition.AppendAtEnd()));
    ExecutorTestPlanSupport.PendingMutation setActiveSheet =
        mutate(new SheetSelector.ByName("Budget Copy"), new MutationAction.SetActiveSheet());
    ExecutorTestPlanSupport.PendingMutation setSelectedSheets =
        mutate(
            new SheetSelector.ByNames(List.of("Budget", "Budget Copy")),
            new MutationAction.SetSelectedSheets());
    ExecutorTestPlanSupport.PendingMutation setSheetVisibility =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN));
    ExecutorTestPlanSupport.PendingMutation setSheetProtection =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetProtection(protectionSettings()));
    ExecutorTestPlanSupport.PendingMutation clearSheetProtection =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearSheetProtection());

    assertNull(formulaFor(copySheet, exception));
    assertNull(formulaFor(setActiveSheet, exception));
    assertNull(formulaFor(setSelectedSheets, exception));
    assertNull(formulaFor(setSheetVisibility, exception));
    assertNull(formulaFor(setSheetProtection, exception));
    assertNull(formulaFor(clearSheetProtection, exception));

    assertEquals("Budget", sheetNameFor(copySheet, exception));
    assertEquals("Budget Copy", sheetNameFor(setActiveSheet, exception));
    assertNull(sheetNameFor(setSelectedSheets, exception));
    assertEquals("Budget", sheetNameFor(setSheetVisibility, exception));
    assertEquals("Budget", sheetNameFor(setSheetProtection, exception));
    assertEquals("Budget", sheetNameFor(clearSheetProtection, exception));

    assertNull(addressFor(copySheet, exception));
    assertNull(addressFor(setActiveSheet, exception));
    assertNull(addressFor(setSelectedSheets, exception));
    assertNull(addressFor(setSheetVisibility, exception));
    assertNull(addressFor(setSheetProtection, exception));
    assertNull(addressFor(clearSheetProtection, exception));

    assertNull(rangeFor(copySheet, exception));
    assertNull(rangeFor(setActiveSheet, exception));
    assertNull(rangeFor(setSelectedSheets, exception));
    assertNull(rangeFor(setSheetVisibility, exception));
    assertNull(rangeFor(setSheetProtection, exception));
    assertNull(rangeFor(clearSheetProtection, exception));

    assertNull(namedRangeNameFor(copySheet, exception));
    assertNull(namedRangeNameFor(setActiveSheet, exception));
    assertNull(namedRangeNameFor(setSelectedSheets, exception));
    assertNull(namedRangeNameFor(setSheetVisibility, exception));
    assertNull(namedRangeNameFor(setSheetProtection, exception));
    assertNull(namedRangeNameFor(clearSheetProtection, exception));
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void extractsContextForStructuralLayoutOperations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation mergeCells =
        mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.MergeCells());
    ExecutorTestPlanSupport.PendingMutation unmergeCells =
        mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.UnmergeCells());
    ExecutorTestPlanSupport.PendingMutation setColumnWidth =
        mutate(
            new ColumnBandSelector.Span("Budget", 0, 1), new MutationAction.SetColumnWidth(16.0));
    ExecutorTestPlanSupport.PendingMutation setRowHeight =
        mutate(new RowBandSelector.Span("Budget", 0, 1), new MutationAction.SetRowHeight(28.5));
    ExecutorTestPlanSupport.PendingMutation insertRows =
        mutate(new RowBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertRows());
    ExecutorTestPlanSupport.PendingMutation deleteRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteRows());
    ExecutorTestPlanSupport.PendingMutation shiftRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftRows(1));
    ExecutorTestPlanSupport.PendingMutation insertColumns =
        mutate(
            new ColumnBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertColumns());
    ExecutorTestPlanSupport.PendingMutation deleteColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteColumns());
    ExecutorTestPlanSupport.PendingMutation shiftColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftColumns(-1));
    ExecutorTestPlanSupport.PendingMutation setRowVisibility =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.SetRowVisibility(true));
    ExecutorTestPlanSupport.PendingMutation setColumnVisibility =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new MutationAction.SetColumnVisibility(false));
    ExecutorTestPlanSupport.PendingMutation groupRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.GroupRows(true));
    ExecutorTestPlanSupport.PendingMutation ungroupRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupRows());
    ExecutorTestPlanSupport.PendingMutation groupColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.GroupColumns(true));
    ExecutorTestPlanSupport.PendingMutation ungroupColumns =
        mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupColumns());
    ExecutorTestPlanSupport.PendingMutation setSheetPane =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1)));
    ExecutorTestPlanSupport.PendingMutation setSheetZoom =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetSheetZoom(125));
    ExecutorTestPlanSupport.PendingMutation setSheetPresentation =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetSheetPresentation(
                new SheetPresentationInput(
                    new SheetDisplayInput(false, false, false, true, true),
                    ColorInput.rgb("#112233"),
                    new SheetOutlineSummaryInput(false, false),
                    new SheetDefaultsInput(11, 18.5d),
                    List.of(
                        new IgnoredErrorInput(
                            "A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))));
    ExecutorTestPlanSupport.PendingMutation setPrintLayout =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetPrintLayout(
                new PrintLayoutInput(
                    new PrintAreaInput.Range("A1:B12"),
                    ExcelPrintOrientation.LANDSCAPE,
                    new PrintScalingInput.Fit(1, 0),
                    new PrintTitleRowsInput.Band(0, 0),
                    new PrintTitleColumnsInput.Band(0, 0),
                    headerFooter("Budget", "", ""),
                    headerFooter("", "Page &P", ""))));
    ExecutorTestPlanSupport.PendingMutation clearPrintLayout =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearPrintLayout());

    assertNull(formulaFor(mergeCells, exception));
    assertEquals("Budget", sheetNameFor(mergeCells, exception));
    assertNull(addressFor(mergeCells, exception));
    assertEquals("A1:B2", rangeFor(mergeCells, exception));

    assertNull(formulaFor(unmergeCells, exception));
    assertEquals("Budget", sheetNameFor(unmergeCells, exception));
    assertNull(addressFor(unmergeCells, exception));
    assertEquals("A1:B2", rangeFor(unmergeCells, exception));

    assertNull(formulaFor(setColumnWidth, exception));
    assertEquals("Budget", sheetNameFor(setColumnWidth, exception));
    assertNull(addressFor(setColumnWidth, exception));
    assertNull(rangeFor(setColumnWidth, exception));

    assertNull(formulaFor(setRowHeight, exception));
    assertEquals("Budget", sheetNameFor(setRowHeight, exception));
    assertNull(addressFor(setRowHeight, exception));
    assertNull(rangeFor(setRowHeight, exception));

    assertNull(formulaFor(insertRows, exception));
    assertEquals("Budget", sheetNameFor(insertRows, exception));
    assertNull(addressFor(insertRows, exception));
    assertNull(rangeFor(insertRows, exception));

    assertNull(formulaFor(deleteRows, exception));
    assertEquals("Budget", sheetNameFor(deleteRows, exception));
    assertNull(addressFor(deleteRows, exception));
    assertNull(rangeFor(deleteRows, exception));

    assertNull(formulaFor(shiftRows, exception));
    assertEquals("Budget", sheetNameFor(shiftRows, exception));
    assertNull(addressFor(shiftRows, exception));
    assertNull(rangeFor(shiftRows, exception));

    assertNull(formulaFor(insertColumns, exception));
    assertEquals("Budget", sheetNameFor(insertColumns, exception));
    assertNull(addressFor(insertColumns, exception));
    assertNull(rangeFor(insertColumns, exception));

    assertNull(formulaFor(deleteColumns, exception));
    assertEquals("Budget", sheetNameFor(deleteColumns, exception));
    assertNull(addressFor(deleteColumns, exception));
    assertNull(rangeFor(deleteColumns, exception));

    assertNull(formulaFor(shiftColumns, exception));
    assertEquals("Budget", sheetNameFor(shiftColumns, exception));
    assertNull(addressFor(shiftColumns, exception));
    assertNull(rangeFor(shiftColumns, exception));

    assertNull(formulaFor(setRowVisibility, exception));
    assertEquals("Budget", sheetNameFor(setRowVisibility, exception));
    assertNull(addressFor(setRowVisibility, exception));
    assertNull(rangeFor(setRowVisibility, exception));

    assertNull(formulaFor(setColumnVisibility, exception));
    assertEquals("Budget", sheetNameFor(setColumnVisibility, exception));
    assertNull(addressFor(setColumnVisibility, exception));
    assertNull(rangeFor(setColumnVisibility, exception));

    assertNull(formulaFor(groupRows, exception));
    assertEquals("Budget", sheetNameFor(groupRows, exception));
    assertNull(addressFor(groupRows, exception));
    assertNull(rangeFor(groupRows, exception));

    assertNull(formulaFor(ungroupRows, exception));
    assertEquals("Budget", sheetNameFor(ungroupRows, exception));
    assertNull(addressFor(ungroupRows, exception));
    assertNull(rangeFor(ungroupRows, exception));

    assertNull(formulaFor(groupColumns, exception));
    assertEquals("Budget", sheetNameFor(groupColumns, exception));
    assertNull(addressFor(groupColumns, exception));
    assertNull(rangeFor(groupColumns, exception));

    assertNull(formulaFor(ungroupColumns, exception));
    assertEquals("Budget", sheetNameFor(ungroupColumns, exception));
    assertNull(addressFor(ungroupColumns, exception));
    assertNull(rangeFor(ungroupColumns, exception));

    assertNull(formulaFor(setSheetPane, exception));
    assertEquals("Budget", sheetNameFor(setSheetPane, exception));
    assertNull(addressFor(setSheetPane, exception));
    assertNull(rangeFor(setSheetPane, exception));

    assertNull(formulaFor(setSheetZoom, exception));
    assertEquals("Budget", sheetNameFor(setSheetZoom, exception));
    assertNull(addressFor(setSheetZoom, exception));
    assertNull(rangeFor(setSheetZoom, exception));

    assertNull(formulaFor(setSheetPresentation, exception));
    assertEquals("Budget", sheetNameFor(setSheetPresentation, exception));
    assertNull(addressFor(setSheetPresentation, exception));
    assertNull(rangeFor(setSheetPresentation, exception));

    assertNull(formulaFor(setPrintLayout, exception));
    assertEquals("Budget", sheetNameFor(setPrintLayout, exception));
    assertNull(addressFor(setPrintLayout, exception));
    assertNull(rangeFor(setPrintLayout, exception));

    assertNull(formulaFor(clearPrintLayout, exception));
    assertEquals("Budget", sheetNameFor(clearPrintLayout, exception));
    assertNull(addressFor(clearPrintLayout, exception));
    assertNull(rangeFor(clearPrintLayout, exception));
  }

  @Test
  void convertsWaveThreeOperationsIntoWorkbookCommands() {
    WorkbookCommand setHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))));
    WorkbookCommand clearHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()));
    WorkbookCommand setComment =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(text("Review"), "GridGrind", false))));
    WorkbookCommand clearComment =
        command(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()));
    WorkbookCommand setNamedRange =
        command(
            mutate(
                new MutationAction.SetNamedRange(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Budget", "B4"))));
    WorkbookCommand deleteNamedRange =
        command(
            mutate(
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                    "BudgetTotal", "Budget"),
                new MutationAction.DeleteNamedRange()));

    assertInstanceOf(WorkbookCommand.SetHyperlink.class, setHyperlink);
    assertInstanceOf(WorkbookCommand.ClearHyperlink.class, clearHyperlink);
    assertInstanceOf(WorkbookCommand.SetComment.class, setComment);
    assertInstanceOf(WorkbookCommand.ClearComment.class, clearComment);
    assertInstanceOf(WorkbookCommand.SetNamedRange.class, setNamedRange);
    assertInstanceOf(WorkbookCommand.DeleteNamedRange.class, deleteNamedRange);
    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"),
        cast(WorkbookCommand.SetHyperlink.class, setHyperlink).target());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        cast(WorkbookCommand.SetComment.class, setComment).comment());
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("Budget", "B4")),
        cast(WorkbookCommand.SetNamedRange.class, setNamedRange).definition());
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        cast(WorkbookCommand.DeleteNamedRange.class, deleteNamedRange).scope());
  }

  @Test
  void convertsRemainingWorkbookOperationsIntoWorkbookCommands() {
    assertInstanceOf(
        WorkbookCommand.CreateSheet.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet())));
    assertInstanceOf(
        WorkbookCommand.RenameSheet.class,
        command(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.RenameSheet("Summary"))));
    assertInstanceOf(
        WorkbookCommand.DeleteSheet.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.DeleteSheet())));
    assertInstanceOf(
        WorkbookCommand.MoveSheet.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.MoveSheet(0))));
    assertInstanceOf(
        WorkbookCommand.MergeCells.class,
        command(
            mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.MergeCells())));
    assertInstanceOf(
        WorkbookCommand.UnmergeCells.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.UnmergeCells())));
    assertInstanceOf(
        WorkbookCommand.SetColumnWidth.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 0, 1),
                new MutationAction.SetColumnWidth(16.0))));
    assertInstanceOf(
        WorkbookCommand.SetRowHeight.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 0, 1), new MutationAction.SetRowHeight(28.5))));
    assertInstanceOf(
        WorkbookCommand.InsertRows.class,
        command(
            mutate(
                new RowBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertRows())));
    assertInstanceOf(
        WorkbookCommand.DeleteRows.class,
        command(mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteRows())));
    assertInstanceOf(
        WorkbookCommand.ShiftRows.class,
        command(mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftRows(1))));
    assertInstanceOf(
        WorkbookCommand.InsertColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Insertion("Budget", 1, 2),
                new MutationAction.InsertColumns())));
    assertInstanceOf(
        WorkbookCommand.DeleteColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteColumns())));
    assertInstanceOf(
        WorkbookCommand.ShiftColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftColumns(-1))));
    assertInstanceOf(
        WorkbookCommand.SetRowVisibility.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetRowVisibility(true))));
    assertInstanceOf(
        WorkbookCommand.SetColumnVisibility.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetColumnVisibility(true))));
    assertInstanceOf(
        WorkbookCommand.GroupRows.class,
        command(
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.GroupRows(false))));
    assertInstanceOf(
        WorkbookCommand.UngroupRows.class,
        command(
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupRows())));
    assertInstanceOf(
        WorkbookCommand.GroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new MutationAction.GroupColumns(false))));
    assertInstanceOf(
        WorkbookCommand.UngroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupColumns())));
    assertInstanceOf(
        WorkbookCommand.SetSheetPane.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1)))));
    assertInstanceOf(
        WorkbookCommand.SetSheetZoom.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetSheetZoom(125))));
    assertInstanceOf(
        WorkbookCommand.SetSheetPresentation.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPresentation(
                    new SheetPresentationInput(
                        new SheetDisplayInput(false, false, false, true, true),
                        ColorInput.rgb("#112233"),
                        new SheetOutlineSummaryInput(false, false),
                        new SheetDefaultsInput(11, 18.5d),
                        List.of(
                            new IgnoredErrorInput(
                                "A1:B2",
                                List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))))));
    assertInstanceOf(
        WorkbookCommand.SetPrintLayout.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetPrintLayout(
                    new PrintLayoutInput(
                        new PrintAreaInput.Range("A1:B12"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingInput.Fit(1, 0),
                        new PrintTitleRowsInput.Band(0, 0),
                        new PrintTitleColumnsInput.Band(0, 0),
                        headerFooter("Budget", "", ""),
                        headerFooter("", "Page &P", ""))))));
    assertInstanceOf(
        WorkbookCommand.ClearPrintLayout.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearPrintLayout())));
    assertInstanceOf(
        WorkbookCommand.SetCell.class,
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetCell(textCell("x")))));
    assertInstanceOf(
        WorkbookCommand.SetRange.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new MutationAction.SetRange(List.of(List.of(textCell("x")))))));
    assertInstanceOf(
        WorkbookCommand.ClearRange.class,
        command(
            mutate(new RangeSelector.ByRange("Budget", "A1:B1"), new MutationAction.ClearRange())));
    assertInstanceOf(
        WorkbookCommand.SetConditionalFormatting.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("A1:A3"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "A1>0",
                                false,
                                new DifferentialStyleInput(
                                    null,
                                    true,
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    java.util.Optional.empty()))))))));
    assertInstanceOf(
        WorkbookCommand.ClearConditionalFormatting.class,
        command(
            mutate(
                new RangeSelector.AllOnSheet("Budget"),
                new MutationAction.ClearConditionalFormatting())));
    assertInstanceOf(
        WorkbookCommand.SetAutofilter.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B4"), new MutationAction.SetAutofilter())));
    assertInstanceOf(
        WorkbookCommand.ClearAutofilter.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter())));
    assertInstanceOf(
        WorkbookCommand.SetTable.class,
        command(
            mutate(
                new MutationAction.SetTable(
                    new TableInput(
                        "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None())))));
    assertInstanceOf(
        WorkbookCommand.DeleteTable.class,
        command(
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new MutationAction.DeleteTable())));
    assertInstanceOf(
        WorkbookCommand.AppendRow.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.AppendRow(List.of(textCell("x"))))));
    assertInstanceOf(
        WorkbookCommand.AutoSizeColumns.class,
        command(mutate(new SheetSelector.ByName("Budget"), new MutationAction.AutoSizeColumns())));

    WorkbookCommand.SetSheetPane setSheetPaneNone =
        cast(
            WorkbookCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetPane(new PaneInput.None()))));
    WorkbookCommand.SetSheetPane setSheetPaneSplit =
        cast(
            WorkbookCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetPane(
                        new PaneInput.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT)))));
    WorkbookCommand.SetPrintLayout defaultPrintLayout =
        cast(
            WorkbookCommand.SetPrintLayout.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetPrintLayout(PrintLayoutInput.defaults()))));
    WorkbookCommand.SetSheetPresentation defaultSheetPresentation =
        cast(
            WorkbookCommand.SetSheetPresentation.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetPresentation(SheetPresentationInput.defaults()))));

    assertEquals(new ExcelSheetPane.None(), setSheetPaneNone.pane());
    assertEquals(
        new ExcelSheetPane.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
        setSheetPaneSplit.pane());
    assertEquals(ExcelSheetPresentation.defaults(), defaultSheetPresentation.presentation());
    assertEquals(
        new dev.erst.gridgrind.excel.ExcelPrintLayout(
            new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
            ExcelPrintOrientation.PORTRAIT,
            new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")),
        defaultPrintLayout.printLayout());
  }

  @Test
  void convertsStructuralAndSurfaceReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand workbookSummary =
        readCommand(
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()));
    WorkbookReadCommand workbookProtection =
        readCommand(
            inspect(
                "workbook-protection",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookProtection()));
    WorkbookReadCommand namedRanges =
        readCommand(
            inspect(
                "ranges",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName(
                            "BudgetTotal"),
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                            "LocalItem", "Budget"))),
                new InspectionQuery.GetNamedRanges()));
    WorkbookReadCommand sheetSummary =
        readCommand(
            inspect(
                "sheet",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetSheetSummary()));
    WorkbookReadCommand cells =
        readCommand(
            inspect(
                "cells",
                new CellSelector.ByAddresses("Budget", List.of("A1")),
                new InspectionQuery.GetCells()));
    WorkbookReadCommand window =
        readCommand(
            inspect(
                "window",
                new RangeSelector.RectangularWindow("Budget", "A1", 2, 2),
                new InspectionQuery.GetWindow()));
    WorkbookReadCommand merged =
        readCommand(
            inspect(
                "merged",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetMergedRegions()));
    WorkbookReadCommand hyperlinks =
        readCommand(
            inspect(
                "hyperlinks",
                new CellSelector.AllUsedInSheet("Budget"),
                new InspectionQuery.GetHyperlinks()));
    WorkbookReadCommand comments =
        readCommand(
            inspect(
                "comments",
                new CellSelector.ByAddresses("Budget", List.of("A1")),
                new InspectionQuery.GetComments()));
    WorkbookReadCommand layout =
        readCommand(
            inspect(
                "layout",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetSheetLayout()));
    WorkbookReadCommand printLayout =
        readCommand(
            inspect(
                "printLayout",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetPrintLayout()));
    WorkbookReadCommand validations =
        readCommand(
            inspect(
                "validations",
                new RangeSelector.AllOnSheet("Budget"),
                new InspectionQuery.GetDataValidations()));
    WorkbookReadCommand conditionalFormatting =
        readCommand(
            inspect(
                "conditionalFormatting",
                new RangeSelector.ByRanges("Budget", List.of("A1:A3")),
                new InspectionQuery.GetConditionalFormatting()));
    WorkbookReadCommand autofilters =
        readCommand(
            inspect(
                "autofilters",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetAutofilters()));
    WorkbookReadCommand tables =
        readCommand(
            inspect(
                "tables",
                new TableSelector.ByNames(List.of("BudgetTable")),
                new InspectionQuery.GetTables()));
    WorkbookReadCommand formulaSurface =
        readCommand(
            inspect(
                "formula",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.GetFormulaSurface()));
    WorkbookReadCommand schema =
        readCommand(
            inspect(
                "schema",
                new RangeSelector.RectangularWindow("Budget", "A1", 3, 2),
                new InspectionQuery.GetSheetSchema()));
    WorkbookReadCommand namedRangeSurface =
        readCommand(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.GetNamedRangeSurface()));

    assertInstanceOf(WorkbookReadCommand.GetWorkbookSummary.class, workbookSummary);
    assertInstanceOf(WorkbookReadCommand.GetWorkbookProtection.class, workbookProtection);
    assertInstanceOf(WorkbookReadCommand.GetNamedRanges.class, namedRanges);
    assertInstanceOf(WorkbookReadCommand.GetSheetSummary.class, sheetSummary);
    assertInstanceOf(WorkbookReadCommand.GetCells.class, cells);
    assertInstanceOf(WorkbookReadCommand.GetWindow.class, window);
    assertInstanceOf(WorkbookReadCommand.GetMergedRegions.class, merged);
    assertInstanceOf(WorkbookReadCommand.GetHyperlinks.class, hyperlinks);
    assertInstanceOf(WorkbookReadCommand.GetComments.class, comments);
    assertInstanceOf(WorkbookReadCommand.GetSheetLayout.class, layout);
    assertInstanceOf(WorkbookReadCommand.GetPrintLayout.class, printLayout);
    assertInstanceOf(WorkbookReadCommand.GetDataValidations.class, validations);
    assertInstanceOf(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting);
    assertInstanceOf(WorkbookReadCommand.GetAutofilters.class, autofilters);
    assertInstanceOf(WorkbookReadCommand.GetTables.class, tables);
    assertInstanceOf(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface);
    assertInstanceOf(WorkbookReadCommand.GetSheetSchema.class, schema);
    assertInstanceOf(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetNamedRanges.class, namedRanges).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.AllUsedCells.class,
        cast(WorkbookReadCommand.GetHyperlinks.class, hyperlinks).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.Selected.class,
        cast(WorkbookReadCommand.GetComments.class, comments).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.All.class,
        cast(WorkbookReadCommand.GetDataValidations.class, validations).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting)
            .selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelTableSelection.ByNames.class,
        cast(WorkbookReadCommand.GetTables.class, tables).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.All.class,
        cast(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface).selection());
  }

  @Test
  void convertsAnalysisReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand formulaHealth =
        readCommand(
            inspect(
                "formulaHealth",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeFormulaHealth()));
    WorkbookReadCommand validationHealth =
        readCommand(
            inspect(
                "validationHealth",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeDataValidationHealth()));
    WorkbookReadCommand conditionalFormattingHealth =
        readCommand(
            inspect(
                "conditionalFormattingHealth",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.AnalyzeConditionalFormattingHealth()));
    WorkbookReadCommand autofilterHealth =
        readCommand(
            inspect(
                "autofilterHealth",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.AnalyzeAutofilterHealth()));
    WorkbookReadCommand tableHealth =
        readCommand(
            inspect(
                "tableHealth",
                new TableSelector.ByNames(List.of("BudgetTable")),
                new InspectionQuery.AnalyzeTableHealth()));
    WorkbookReadCommand hyperlinkHealth =
        readCommand(
            inspect(
                "hyperlinkHealth",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeHyperlinkHealth()));
    WorkbookReadCommand namedRangeHealth =
        readCommand(
            inspect(
                "namedRangeHealth",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.AnalyzeNamedRangeHealth()));
    WorkbookReadCommand workbookFindings =
        readCommand(
            inspect(
                "workbookFindings",
                new WorkbookSelector.Current(),
                new InspectionQuery.AnalyzeWorkbookFindings()));

    assertInstanceOf(WorkbookReadCommand.AnalyzeFormulaHealth.class, formulaHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeDataValidationHealth.class, validationHealth);
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class, conditionalFormattingHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeAutofilterHealth.class, autofilterHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeTableHealth.class, tableHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeHyperlinkHealth.class, hyperlinkHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeNamedRangeHealth.class, namedRangeHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeWorkbookFindings.class, workbookFindings);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(
                WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class,
                conditionalFormattingHealth)
            .selection());
  }

  @Test
  void convertsSheetStateOperationsIntoWorkbookCommandsWithExactFields() {
    WorkbookCommand.CopySheet copySheet =
        cast(
            WorkbookCommand.CopySheet.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.CopySheet(
                        "Budget Copy", new SheetCopyPosition.AtIndex(1)))));
    WorkbookCommand.SetActiveSheet setActiveSheet =
        cast(
            WorkbookCommand.SetActiveSheet.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget Copy"), new MutationAction.SetActiveSheet())));
    WorkbookCommand.SetSelectedSheets setSelectedSheets =
        cast(
            WorkbookCommand.SetSelectedSheets.class,
            command(
                mutate(
                    new SheetSelector.ByNames(List.of("Budget", "Budget Copy")),
                    new MutationAction.SetSelectedSheets())));
    WorkbookCommand.SetSheetVisibility setSheetVisibility =
        cast(
            WorkbookCommand.SetSheetVisibility.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN))));
    WorkbookCommand.SetSheetProtection setSheetProtection =
        cast(
            WorkbookCommand.SetSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.SetSheetProtection(protectionSettings()))));
    WorkbookCommand.ClearSheetProtection clearSheetProtection =
        cast(
            WorkbookCommand.ClearSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new MutationAction.ClearSheetProtection())));

    ExcelSheetCopyPosition.AtIndex position =
        assertInstanceOf(ExcelSheetCopyPosition.AtIndex.class, copySheet.position());

    assertEquals("Budget", copySheet.sourceSheetName());
    assertEquals("Budget Copy", copySheet.newSheetName());
    assertEquals(1, position.targetIndex());
    assertEquals("Budget Copy", setActiveSheet.sheetName());
    assertEquals(List.of("Budget", "Budget Copy"), setSelectedSheets.sheetNames());
    assertEquals(ExcelSheetVisibility.HIDDEN, setSheetVisibility.visibility());
    assertEquals(excelProtectionSettings(), setSheetProtection.protection());
    assertEquals("Budget", clearSheetProtection.sheetName());
  }

  @Test
  void convertsReadResultsIntoProtocolReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    InspectionResult workbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets(
                    1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)));
    InspectionResult namedRanges =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult(
                "ranges",
                List.of(
                    new ExcelNamedRangeSnapshot.RangeSnapshot(
                        "BudgetTotal",
                        new ExcelNamedRangeScope.WorkbookScope(),
                        "Budget!$B$4",
                        new ExcelNamedRangeTarget("Budget", "B4")))));
    InspectionResult sheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VISIBLE,
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Unprotected(),
                    4,
                    3,
                    2)));
    InspectionResult cells =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                "cells", "Budget", List.of(blank)));
    InspectionResult window =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult(
                "window",
                new dev.erst.gridgrind.excel.WorkbookReadResult.Window(
                    "Budget",
                    "A1",
                    1,
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.WindowRow(
                            0, List.of(blank))))));
    InspectionResult merged =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult(
                "merged",
                "Budget",
                List.of(new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegion("A1:B2"))));
    InspectionResult hyperlinks =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult(
                "hyperlinks",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellHyperlink(
                        "A1", new ExcelHyperlink.Url("https://example.com/report")))));
    InspectionResult comments =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult(
                "comments",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellComment(
                        "A1",
                        new ExcelCommentSnapshot("Review", "GridGrind", false, null, null)))));
    InspectionResult layout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                "layout",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelSheetPane.Frozen(1, 1, 1, 1),
                    125,
                    defaultSheetPresentationSnapshot(),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.ColumnLayout(
                            0, 12.5, false, 0, false)),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.RowLayout(
                            0, 18.0, false, 0, false)))));
    InspectionResult printLayout =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                "printLayout",
                "Budget",
                new dev.erst.gridgrind.excel.ExcelPrintLayoutSnapshot(
                    new dev.erst.gridgrind.excel.ExcelPrintLayout(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.Range("A1:B20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Fit(1, 0),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.Band(0, 0),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.Band(0, 0),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("Budget", "", ""),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "Page &P", "")),
                    defaultPrintSetupSnapshot())));
    InspectionResult conditionalFormatting =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult(
                "conditional-formatting",
                "Budget",
                List.of(
                    new ExcelConditionalFormattingBlockSnapshot(
                        List.of("A2:A5"),
                        List.of(
                            new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                                1,
                                true,
                                "A2>0",
                                new ExcelDifferentialStyleSnapshot(
                                    "0.00", true, null, null, "#102030", null, null, "#E0F0AA",
                                    null, List.of())),
                            new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                                2,
                                false,
                                List.of(
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                                List.of("#AA0000", "#00AA00")))))));
    InspectionResult formulaSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult(
                "formula",
                new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface(
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SheetFormulaSurface(
                            "Budget",
                            1,
                            1,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaPattern(
                                    "SUM(B2:B3)", 1, List.of("B4"))))))));
    InspectionResult schema =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult(
                "schema",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema(
                    "Budget",
                    "A1",
                    3,
                    2,
                    2,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SchemaColumn(
                            0,
                            "A1",
                            "Item",
                            2,
                            0,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookReadResult.TypeCount(
                                    "STRING", 2)),
                            "STRING")))));
    InspectionResult namedRangeSurface =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult(
                "surface",
                new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface(
                    1,
                    0,
                    1,
                    0,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceEntry(
                            "BudgetTotal",
                            new ExcelNamedRangeScope.WorkbookScope(),
                            "Budget!$B$4",
                            dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeBackingKind
                                .RANGE)))));
    InspectionResult formulaHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult(
                "formula-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.FormulaHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .FORMULA_VOLATILE_FUNCTION,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.INFO,
                            "Volatile formula",
                            "Formula uses NOW().",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Cell(
                                "Budget", "B4"),
                            List.of("NOW()"))))));

    assertInstanceOf(InspectionResult.WorkbookSummaryResult.class, workbookSummary);
    assertInstanceOf(InspectionResult.NamedRangesResult.class, namedRanges);
    assertInstanceOf(InspectionResult.SheetSummaryResult.class, sheetSummary);
    assertInstanceOf(InspectionResult.CellsResult.class, cells);
    assertInstanceOf(InspectionResult.WindowResult.class, window);
    assertInstanceOf(InspectionResult.MergedRegionsResult.class, merged);
    assertInstanceOf(InspectionResult.HyperlinksResult.class, hyperlinks);
    assertInstanceOf(InspectionResult.CommentsResult.class, comments);
    assertInstanceOf(InspectionResult.SheetLayoutResult.class, layout);
    assertInstanceOf(InspectionResult.PrintLayoutResult.class, printLayout);
    assertInstanceOf(InspectionResult.ConditionalFormattingResult.class, conditionalFormatting);
    assertInstanceOf(InspectionResult.FormulaSurfaceResult.class, formulaSurface);
    assertInstanceOf(InspectionResult.SheetSchemaResult.class, schema);
    assertInstanceOf(InspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface);
    assertInstanceOf(InspectionResult.FormulaHealthResult.class, formulaHealth);
    assertEquals(
        "Budget",
        cast(InspectionResult.WorkbookSummaryResult.class, workbookSummary)
            .workbook()
            .sheetNames()
            .getFirst());
    assertEquals(
        "BudgetTotal",
        cast(InspectionResult.NamedRangesResult.class, namedRanges)
            .namedRanges()
            .getFirst()
            .name());
    assertEquals(
        "Budget",
        cast(InspectionResult.SheetSummaryResult.class, sheetSummary).sheet().sheetName());
    assertEquals(
        "A1", cast(InspectionResult.CellsResult.class, cells).cells().getFirst().address());
    assertEquals(
        "A1",
        cast(InspectionResult.WindowResult.class, window)
            .window()
            .rows()
            .getFirst()
            .cells()
            .getFirst()
            .address());
    assertEquals(
        "A1:B2",
        cast(InspectionResult.MergedRegionsResult.class, merged)
            .mergedRegions()
            .getFirst()
            .range());
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        cast(InspectionResult.HyperlinksResult.class, hyperlinks)
            .hyperlinks()
            .getFirst()
            .hyperlink());
    assertEquals(
        "Review",
        cast(InspectionResult.CommentsResult.class, comments)
            .comments()
            .getFirst()
            .comment()
            .text());
    assertInstanceOf(
        PaneReport.Frozen.class,
        cast(InspectionResult.SheetLayoutResult.class, layout).layout().pane());
    assertEquals(
        125, cast(InspectionResult.SheetLayoutResult.class, layout).layout().zoomPercent());
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        cast(InspectionResult.PrintLayoutResult.class, printLayout).layout().orientation());
    assertEquals(
        2,
        cast(InspectionResult.ConditionalFormattingResult.class, conditionalFormatting)
            .conditionalFormattingBlocks()
            .getFirst()
            .rules()
            .size());
    assertEquals(
        1,
        cast(InspectionResult.FormulaSurfaceResult.class, formulaSurface)
            .analysis()
            .totalFormulaCellCount());
    assertEquals(
        "STRING",
        cast(InspectionResult.SheetSchemaResult.class, schema)
            .analysis()
            .columns()
            .getFirst()
            .dominantType());
    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.RANGE,
        cast(InspectionResult.NamedRangeSurfaceResult.class, namedRangeSurface)
            .analysis()
            .namedRanges()
            .getFirst()
            .kind());
    assertEquals(
        1,
        cast(InspectionResult.FormulaHealthResult.class, formulaHealth)
            .analysis()
            .summary()
            .infoCount());
  }

  @Test
  void convertsRemainingAnalysisReadResultsIntoProtocolReadResults() {
    InspectionResult conditionalFormattingHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult(
                "conditional-formatting-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.ConditionalFormattingHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 1, 0, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
                            "Priority collision",
                            "Conditional-formatting priorities collide.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("FORMULA_RULE@Budget!A1:A3"))))));
    InspectionResult hyperlinkHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult(
                "hyperlink-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.HyperlinkHealth(
                    2,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(2, 1, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .HYPERLINK_MALFORMED_TARGET,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.ERROR,
                            "Malformed hyperlink target",
                            "Hyperlink target is malformed.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .Workbook(),
                            List.of("mailto:")),
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .HYPERLINK_MISSING_DOCUMENT_SHEET,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.WARNING,
                            "Missing target sheet",
                            "Sheet is missing.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("Missing!A1"))))));
    InspectionResult namedRangeHealth =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult(
                "named-range-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.NamedRangeHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .NAMED_RANGE_BROKEN_REFERENCE,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.WARNING,
                            "Broken named range",
                            "Named range contains #REF!.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Range(
                                "Budget", "A1:B2"),
                            List.of("#REF!"))))));
    InspectionResult workbookFindings =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult(
                "workbook-findings",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.WorkbookFindings(
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.foundation.AnalysisFindingCode
                                .NAMED_RANGE_SCOPE_SHADOWING,
                            dev.erst.gridgrind.excel.foundation.AnalysisSeverity.INFO,
                            "Scope shadowing",
                            "Name exists in more than one scope.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .NamedRange(
                                "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
                            List.of("Budget!$B$4"))))));

    assertEquals(
        1,
        cast(InspectionResult.ConditionalFormattingHealthResult.class, conditionalFormattingHealth)
            .analysis()
            .checkedConditionalFormattingBlockCount());
    GridGrindResponse.AnalysisLocationReport workbookLocation =
        cast(InspectionResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindResponse.AnalysisLocationReport sheetLocation =
        cast(InspectionResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .get(1)
            .location();
    GridGrindResponse.AnalysisLocationReport rangeLocation =
        cast(InspectionResult.NamedRangeHealthResult.class, namedRangeHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindResponse.AnalysisLocationReport namedRangeLocation =
        cast(InspectionResult.WorkbookFindingsResult.class, workbookFindings)
            .analysis()
            .findings()
            .getFirst()
            .location();

    assertInstanceOf(GridGrindResponse.AnalysisLocationReport.Workbook.class, workbookLocation);
    assertEquals(
        "Budget",
        cast(GridGrindResponse.AnalysisLocationReport.Sheet.class, sheetLocation).sheetName());
    assertEquals(
        "A1:B2", cast(GridGrindResponse.AnalysisLocationReport.Range.class, rangeLocation).range());
    assertEquals(
        "BudgetTotal",
        cast(GridGrindResponse.AnalysisLocationReport.NamedRange.class, namedRangeLocation).name());
  }

  @Test
  void convertsNamedRangeFormulaSnapshotsAndFormulaBackedSurfaceEntries() {
    GridGrindResponse.NamedRangeReport formulaReport =
        InspectionResultWorkbookCoreReportSupport.toNamedRangeReport(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Budget!$B$2:$B$3)"));
    assertInstanceOf(GridGrindResponse.NamedRangeReport.FormulaReport.class, formulaReport);

    InspectionResult.NamedRangeSurfaceResult surface =
        cast(
            InspectionResult.NamedRangeSurfaceResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult(
                    "surface",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface(
                        0,
                        1,
                        0,
                        1,
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceEntry(
                                "BudgetRollup",
                                new ExcelNamedRangeScope.SheetScope("Budget"),
                                "SUM(Budget!$B$2:$B$3)",
                                dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeBackingKind
                                    .FORMULA))))));

    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.FORMULA,
        surface.analysis().namedRanges().getFirst().kind());
  }

  @Test
  void convertsPaneNoneIntoProtocolReport() {
    InspectionResult.SheetLayoutResult layout =
        cast(
            InspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.None(),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));

    assertInstanceOf(PaneReport.None.class, layout.layout().pane());
  }

  @Test
  void convertsSplitPaneAndDefaultPrintLayoutIntoProtocolReports() {
    InspectionResult.SheetLayoutResult layout =
        cast(
            InspectionResult.SheetLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.Split(
                            1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
                        100,
                        defaultSheetPresentationSnapshot(),
                        List.of(),
                        List.of()))));
    InspectionResult.PrintLayoutResult printLayout =
        cast(
            InspectionResult.PrintLayoutResult.class,
            InspectionResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                    "print-layout",
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelPrintLayoutSnapshot(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout(
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
                            ExcelPrintOrientation.PORTRAIT,
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
                            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
                            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
                            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")),
                        defaultPrintSetupSnapshot()))));

    PaneReport.Split pane = assertInstanceOf(PaneReport.Split.class, layout.layout().pane());
    assertEquals(1200, pane.xSplitPosition());
    assertEquals(2400, pane.ySplitPosition());
    assertEquals(3, pane.leftmostColumn());
    assertEquals(4, pane.topRow());
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, pane.activePane());
    assertInstanceOf(PrintAreaReport.None.class, printLayout.layout().printArea());
    assertInstanceOf(PrintScalingReport.Automatic.class, printLayout.layout().scaling());
    assertInstanceOf(PrintTitleRowsReport.None.class, printLayout.layout().repeatingRows());
    assertInstanceOf(PrintTitleColumnsReport.None.class, printLayout.layout().repeatingColumns());
  }

  @Test
  void convertsSheetStateReadResultsIntoProtocolShapes() {
    InspectionResult emptyWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty(
                    0, List.of(), 0, false)));
    InspectionResult populatedWorkbookSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook-2",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets(
                    2,
                    List.of("Budget", "Archive"),
                    "Archive",
                    List.of("Budget", "Archive"),
                    1,
                    true)));
    InspectionResult protectedSheetSummary =
        InspectionResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VERY_HIDDEN,
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Protected(
                        excelProtectionSettings()),
                    4,
                    7,
                    3)));

    GridGrindResponse.WorkbookSummary.Empty empty =
        assertInstanceOf(
            GridGrindResponse.WorkbookSummary.Empty.class,
            cast(InspectionResult.WorkbookSummaryResult.class, emptyWorkbookSummary).workbook());
    GridGrindResponse.WorkbookSummary.WithSheets populated =
        assertInstanceOf(
            GridGrindResponse.WorkbookSummary.WithSheets.class,
            cast(InspectionResult.WorkbookSummaryResult.class, populatedWorkbookSummary)
                .workbook());
    GridGrindResponse.SheetSummaryReport sheet =
        cast(InspectionResult.SheetSummaryResult.class, protectedSheetSummary).sheet();
    GridGrindResponse.SheetProtectionReport.Protected protection =
        assertInstanceOf(
            GridGrindResponse.SheetProtectionReport.Protected.class, sheet.protection());

    assertEquals(0, empty.sheetCount());
    assertEquals("Archive", populated.activeSheetName());
    assertEquals(List.of("Budget", "Archive"), populated.selectedSheetNames());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, sheet.visibility());
    assertEquals(protectionSettings(), protection.settings());
  }

  @Test
  void readTypeReturnsDiscriminatorsForAllReadVariants() {
    assertEquals(
        List.of(
            "GET_WORKBOOK_SUMMARY",
            "GET_WORKBOOK_PROTECTION",
            "GET_NAMED_RANGES",
            "GET_SHEET_SUMMARY",
            "GET_CELLS",
            "GET_WINDOW",
            "GET_MERGED_REGIONS",
            "GET_HYPERLINKS",
            "GET_COMMENTS",
            "GET_SHEET_LAYOUT",
            "GET_PRINT_LAYOUT",
            "GET_DATA_VALIDATIONS",
            "GET_CONDITIONAL_FORMATTING",
            "GET_AUTOFILTERS",
            "GET_TABLES",
            "GET_FORMULA_SURFACE",
            "GET_SHEET_SCHEMA",
            "GET_NAMED_RANGE_SURFACE",
            "ANALYZE_FORMULA_HEALTH",
            "ANALYZE_DATA_VALIDATION_HEALTH",
            "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
            "ANALYZE_AUTOFILTER_HEALTH",
            "ANALYZE_TABLE_HEALTH",
            "ANALYZE_HYPERLINK_HEALTH",
            "ANALYZE_NAMED_RANGE_HEALTH",
            "ANALYZE_WORKBOOK_FINDINGS"),
        List.of(
            readType(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())),
            readType(
                inspect(
                    "workbook-protection",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookProtection())),
            readType(
                inspect(
                    "ranges",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            readType(
                inspect(
                    "sheet",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetSummary())),
            readType(
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new InspectionQuery.GetCells())),
            readType(
                inspect(
                    "window",
                    new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
                    new InspectionQuery.GetWindow())),
            readType(
                inspect(
                    "merged",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetMergedRegions())),
            readType(
                inspect(
                    "hyperlinks",
                    new CellSelector.AllUsedInSheet("Budget"),
                    new InspectionQuery.GetHyperlinks())),
            readType(
                inspect(
                    "comments",
                    new CellSelector.AllUsedInSheet("Budget"),
                    new InspectionQuery.GetComments())),
            readType(
                inspect(
                    "layout",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetLayout())),
            readType(
                inspect(
                    "print-layout",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetPrintLayout())),
            readType(
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetDataValidations())),
            readType(
                inspect(
                    "conditional-formatting",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetConditionalFormatting())),
            readType(
                inspect(
                    "autofilters",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetAutofilters())),
            readType(
                inspect(
                    "tables",
                    new TableSelector.ByNames(List.of("BudgetTable")),
                    new InspectionQuery.GetTables())),
            readType(
                inspect(
                    "formula", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface())),
            readType(
                inspect(
                    "schema",
                    new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
                    new InspectionQuery.GetSheetSchema())),
            readType(
                inspect(
                    "surface",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRangeSurface())),
            readType(
                inspect(
                    "formula-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeFormulaHealth())),
            readType(
                inspect(
                    "validation-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeDataValidationHealth())),
            readType(
                inspect(
                    "conditional-formatting-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeConditionalFormattingHealth())),
            readType(
                inspect(
                    "autofilter-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeAutofilterHealth())),
            readType(
                inspect(
                    "table-health",
                    new TableSelector.All(),
                    new InspectionQuery.AnalyzeTableHealth())),
            readType(
                inspect(
                    "hyperlink-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeHyperlinkHealth())),
            readType(
                inspect(
                    "named-range-health",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.AnalyzeNamedRangeHealth())),
            readType(
                inspect(
                    "workbook-findings",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.AnalyzeWorkbookFindings()))));
  }

  @Test
  void extractsContextForReadOperations() {
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    RuntimeException runtimeException = new RuntimeException("x");

    InspectionStep workbook =
        inspect(
            "workbook", new WorkbookSelector.Current(), new InspectionQuery.GetWorkbookSummary());
    InspectionStep workbookProtection =
        inspect(
            "workbook-protection",
            new WorkbookSelector.Current(),
            new InspectionQuery.GetWorkbookProtection());
    InspectionStep namedRanges =
        inspect(
            "ranges",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
            new InspectionQuery.GetNamedRanges());
    InspectionStep sheet =
        inspect("sheet", new SheetSelector.ByName("Budget"), new InspectionQuery.GetSheetSummary());
    InspectionStep cells =
        inspect(
            "cells",
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new InspectionQuery.GetCells());
    InspectionStep window =
        inspect(
            "window",
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2),
            new InspectionQuery.GetWindow());
    InspectionStep merged =
        inspect(
            "merged", new SheetSelector.ByName("Budget"), new InspectionQuery.GetMergedRegions());
    InspectionStep hyperlinks =
        inspect(
            "hyperlinks",
            new CellSelector.AllUsedInSheet("Budget"),
            new InspectionQuery.GetHyperlinks());
    InspectionStep comments =
        inspect(
            "comments",
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new InspectionQuery.GetComments());
    InspectionStep layout =
        inspect("layout", new SheetSelector.ByName("Budget"), new InspectionQuery.GetSheetLayout());
    InspectionStep printLayout =
        inspect(
            "print-layout",
            new SheetSelector.ByName("Budget"),
            new InspectionQuery.GetPrintLayout());
    InspectionStep validations =
        inspect(
            "validations",
            new RangeSelector.ByRanges("Budget", List.of("A1:A3")),
            new InspectionQuery.GetDataValidations());
    InspectionStep conditionalFormatting =
        inspect(
            "conditional-formatting",
            new RangeSelector.ByRanges("Budget", List.of("B2:B5")),
            new InspectionQuery.GetConditionalFormatting());
    InspectionStep autofilters =
        inspect(
            "autofilters",
            new SheetSelector.ByName("Budget"),
            new InspectionQuery.GetAutofilters());
    InspectionStep tables =
        inspect(
            "tables",
            new TableSelector.ByNames(List.of("BudgetTable")),
            new InspectionQuery.GetTables());
    InspectionStep formula =
        inspect("formula", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface());
    InspectionStep schema =
        inspect(
            "schema",
            new RangeSelector.RectangularWindow("Budget", "C3", 2, 2),
            new InspectionQuery.GetSheetSchema());
    InspectionStep surface =
        inspect(
            "surface",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
            new InspectionQuery.GetNamedRangeSurface());
    InspectionStep formulaHealth =
        inspect(
            "formula-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeFormulaHealth());
    InspectionStep validationHealth =
        inspect(
            "validation-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeDataValidationHealth());
    InspectionStep conditionalFormattingHealth =
        inspect(
            "conditional-formatting-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeConditionalFormattingHealth());
    InspectionStep autofilterHealth =
        inspect(
            "autofilter-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeAutofilterHealth());
    InspectionStep tableHealth =
        inspect(
            "table-health",
            new TableSelector.ByNames(List.of("BudgetTable")),
            new InspectionQuery.AnalyzeTableHealth());
    InspectionStep hyperlinkHealth =
        inspect(
            "hyperlink-health",
            new SheetSelector.All(),
            new InspectionQuery.AnalyzeHyperlinkHealth());
    InspectionStep namedRangeHealth =
        inspect(
            "named-range-health",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                List.of(
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                        "BudgetTotal"))),
            new InspectionQuery.AnalyzeNamedRangeHealth());
    InspectionStep workbookFindings =
        inspect(
            "workbook-findings",
            new WorkbookSelector.Current(),
            new InspectionQuery.AnalyzeWorkbookFindings());

    assertReadContext(workbook, null, null, null, runtimeException);
    assertReadContext(workbookProtection, null, null, null, runtimeException);
    assertReadContext(namedRanges, null, null, null, runtimeException);
    assertReadContext(sheet, "Budget", null, null, runtimeException);
    assertReadContext(cells, "Budget", null, null, runtimeException);
    assertReadContext(window, "Budget", "B2", null, runtimeException);
    assertReadContext(merged, "Budget", null, null, runtimeException);
    assertReadContext(hyperlinks, "Budget", null, null, runtimeException);
    assertReadContext(comments, "Budget", null, null, runtimeException);
    assertReadContext(layout, "Budget", null, null, runtimeException);
    assertReadContext(printLayout, "Budget", null, null, runtimeException);
    assertReadContext(validations, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormatting, "Budget", null, null, runtimeException);
    assertReadContext(autofilters, "Budget", null, null, runtimeException);
    assertReadContext(tables, null, null, null, runtimeException);
    assertReadContext(formula, null, null, null, runtimeException);
    assertReadContext(schema, "Budget", "C3", null, runtimeException);
    assertReadContext(surface, null, null, null, runtimeException);
    assertReadContext(formulaHealth, "Budget", null, null, runtimeException);
    assertReadContext(validationHealth, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormattingHealth, "Budget", null, null, runtimeException);
    assertReadContext(autofilterHealth, "Budget", null, null, runtimeException);
    assertReadContext(tableHealth, null, null, null, runtimeException);
    assertReadContext(hyperlinkHealth, null, null, null, runtimeException);
    assertReadContext(namedRangeHealth, null, null, "BudgetTotal", runtimeException);
    assertReadContext(workbookFindings, null, null, null, runtimeException);

    assertEquals("BudgetTotal", namedRangeNameFor(workbook, missingNamedRange));
    assertEquals("BudgetTotal", namedRangeNameFor(namedRanges, missingNamedRange));
    assertEquals("BAD!", addressFor(cells, invalidAddress));
  }

  @Test
  void extractsSingleSheetAndNamedRangeContextOnlyWhenSelectionsAreUnambiguous() {
    assertEquals(
        "Budget",
        sheetNameFor(
            inspect(
                "formula",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.GetFormulaSurface())));
    assertNull(
        sheetNameFor(
            inspect(
                "hyperlink-health",
                new SheetSelector.ByNames(List.of("Budget", "Forecast")),
                new InspectionQuery.AnalyzeHyperlinkHealth())));

    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName(
                            "BudgetTotal"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                            "BudgetTotal", "Budget"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
    assertNull(
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "BudgetTotal"),
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "ForecastTotal"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
  }

  @Test
  void extractsReadContextFromExceptionsBeforeFallingBackToReadShape() {
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());

    assertEquals(
        "C3",
        addressFor(
            inspect(
                "schema",
                new RangeSelector.RectangularWindow("Budget", "C3", 2, 2),
                new InspectionQuery.GetSheetSchema()),
            invalidAddress));
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.GetNamedRangeSurface()),
            missingNamedRange));
  }

  @Test
  void extractsContextForAuthoringMetadataAndNamedRangeOperations() {
    RuntimeException exception = new RuntimeException("test");
    assertWriteContext(
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetHyperlink(new HyperlinkTarget.Url("https://example.com/report"))),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetComment(new CommentInput(text("Review"), "GridGrind", false))),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new MutationAction.ApplyStyle(
                new CellStyleInput(
                    null,
                    null,
                    new CellFontInput(true, null, null, null, null, null, null),
                    null,
                    null,
                    null))),
        exception,
        "Budget",
        null,
        "A1:B2",
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.ByRange("Budget", "B2:B5"),
            new MutationAction.SetDataValidation(
                new DataValidationInput(
                    new DataValidationRuleInput.TextLength(
                        ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                    true,
                    false,
                    prompt("Reason", "Use 20 characters or fewer.", true),
                    null))),
        exception,
        "Budget",
        null,
        "B2:B5",
        null);
    assertWriteContext(
        mutate(new RangeSelector.AllOnSheet("Budget"), new MutationAction.ClearDataValidations()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.SetConditionalFormatting(
                new ConditionalFormattingBlockInput(
                    List.of("B2:B5"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "B2>0",
                            true,
                            new DifferentialStyleInput(
                                null,
                                true,
                                null,
                                null,
                                java.util.Optional.empty(),
                                null,
                                null,
                                java.util.Optional.empty(),
                                java.util.Optional.empty())))))),
        exception,
        "Budget",
        null,
        "B2:B5",
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.AllOnSheet("Budget"),
            new MutationAction.ClearConditionalFormatting()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(new RangeSelector.ByRange("Budget", "A1:C4"), new MutationAction.SetAutofilter()),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None()))),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        mutate(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
            new MutationAction.DeleteTable()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new MutationAction.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Budget", "B4"))),
        exception,
        "Budget",
        null,
        "B4",
        "BudgetTotal");
    assertWriteContext(
        mutate(
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                "BudgetTotal"),
            new MutationAction.DeleteNamedRange()),
        exception,
        null,
        null,
        null,
        "BudgetTotal");
    assertWriteContext(
        mutate(
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                "LocalItem", "Budget"),
            new MutationAction.DeleteNamedRange()),
        exception,
        "Budget",
        null,
        null,
        "LocalItem");
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
  }

  @Test
  void extractsAddressAndRangeFallbacksAndExhaustsNamedRangeNullArms() {
    InvalidFormulaException invalidFormula =
        new InvalidFormulaException(
            "Budget", "C3", "SUM(", "bad formula", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    InvalidRangeAddressException invalidRange =
        new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"));

    assertEquals(
        "C3",
        addressFor(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            invalidFormula));
    assertEquals(
        "BAD!",
        addressFor(
            mutate(
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                    "BudgetTotal"),
                new MutationAction.DeleteNamedRange()),
            invalidAddress));

    assertEquals(
        "C1:D2",
        rangeFor(
            mutate(
                new RangeSelector.ByRange("Budget", "C1:D2"),
                new MutationAction.SetRange(List.of(List.of(textCell("x"))))),
            invalidFormula));
    assertEquals(
        "E1:E2",
        rangeFor(
            mutate(new RangeSelector.ByRange("Budget", "E1:E2"), new MutationAction.ClearRange()),
            invalidFormula));
    assertEquals(
        "B2:B5",
        rangeFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("B2:B5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "B2>0",
                                true,
                                new DifferentialStyleInput(
                                    null,
                                    true,
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    java.util.Optional.empty())))))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("B2:B5", "D2:D5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "B2>0",
                                true,
                                new DifferentialStyleInput(
                                    null,
                                    true,
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    java.util.Optional.empty())))))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(text("Review"), "GridGrind", false))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()),
            invalidFormula));
    assertEquals(
        "A1:",
        rangeFor(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            invalidRange));

    List<ExecutorTestPlanSupport.PendingMutation> operationsWithoutNamedRanges =
        List.of(
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.RenameSheet("Summary")),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.DeleteSheet()),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.MoveSheet(0)),
            mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.MergeCells()),
            mutate(new RangeSelector.ByRange("Budget", "A1:B2"), new MutationAction.UnmergeCells()),
            mutate(
                new ColumnBandSelector.Span("Budget", 0, 1),
                new MutationAction.SetColumnWidth(16.0)),
            mutate(new RowBandSelector.Span("Budget", 0, 1), new MutationAction.SetRowHeight(28.5)),
            mutate(new RowBandSelector.Insertion("Budget", 1, 2), new MutationAction.InsertRows()),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteRows()),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftRows(1)),
            mutate(
                new ColumnBandSelector.Insertion("Budget", 1, 2),
                new MutationAction.InsertColumns()),
            mutate(new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.DeleteColumns()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.ShiftColumns(-1)),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetRowVisibility(true)),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new MutationAction.SetColumnVisibility(false)),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.GroupRows(true)),
            mutate(new RowBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupRows()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.GroupColumns(true)),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2), new MutationAction.UngroupColumns()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.SetSheetZoom(125)),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetPrintLayout(
                    new PrintLayoutInput(
                        new PrintAreaInput.Range("A1:B12"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingInput.Fit(1, 0),
                        new PrintTitleRowsInput.Band(0, 0),
                        new PrintTitleColumnsInput.Band(0, 0),
                        headerFooter("Budget", "", ""),
                        headerFooter("", "Page &P", "")))),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearPrintLayout()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetCell(textCell("x"))),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new MutationAction.SetRange(List.of(List.of(textCell("x"))))),
            mutate(new RangeSelector.ByRange("Budget", "A1:B1"), new MutationAction.ClearRange()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))),
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearHyperlink()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new MutationAction.SetComment(
                    new CommentInput(text("Review"), "GridGrind", false))),
            mutate(new CellSelector.ByAddress("Budget", "A1"), new MutationAction.ClearComment()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new MutationAction.ApplyStyle(
                    new CellStyleInput(
                        null,
                        null,
                        new CellFontInput(true, null, null, null, null, null, null),
                        null,
                        null,
                        null))),
            mutate(
                new RangeSelector.ByRange("Budget", "B2:B5"),
                new MutationAction.SetDataValidation(
                    new DataValidationInput(
                        new DataValidationRuleInput.TextLength(
                            ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                        true,
                        false,
                        null,
                        null))),
            mutate(
                new RangeSelector.AllOnSheet("Budget"), new MutationAction.ClearDataValidations()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetConditionalFormatting(
                    new ConditionalFormattingBlockInput(
                        List.of("B2:B5"),
                        List.of(
                            new ConditionalFormattingRuleInput.FormulaRule(
                                "B2>0",
                                true,
                                new DifferentialStyleInput(
                                    null,
                                    true,
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    null,
                                    null,
                                    java.util.Optional.empty(),
                                    java.util.Optional.empty())))))),
            mutate(
                new RangeSelector.AllOnSheet("Budget"),
                new MutationAction.ClearConditionalFormatting()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:C4"), new MutationAction.SetAutofilter()),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.ClearAutofilter()),
            mutate(
                new MutationAction.SetTable(
                    new TableInput(
                        "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None()))),
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new MutationAction.DeleteTable()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.AppendRow(List.of(textCell("x")))),
            mutate(new SheetSelector.ByName("Budget"), new MutationAction.AutoSizeColumns()));

    for (ExecutorTestPlanSupport.PendingMutation operation : operationsWithoutNamedRanges) {
      assertNull(namedRangeNameFor(operation, invalidFormula));
    }

    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new MutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
            missingNamedRange));
  }
}
