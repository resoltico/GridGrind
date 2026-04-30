package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Mutation-context extraction coverage for the request executor. */
class DefaultGridGrindRequestExecutorMutationContextTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void extractsSheetOnlyContextForDeleteSheetOperations() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation ensureSheet =
        mutate(new SheetSelector.ByName("Archive"), new WorkbookMutationAction.EnsureSheet());
    ExecutorTestPlanSupport.PendingMutation deleteSheet =
        mutate(new SheetSelector.ByName("Archive"), new WorkbookMutationAction.DeleteSheet());

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
            new WorkbookMutationAction.CopySheet(
                "Budget Copy", new SheetCopyPosition.AppendAtEnd()));
    ExecutorTestPlanSupport.PendingMutation setActiveSheet =
        mutate(
            new SheetSelector.ByName("Budget Copy"), new WorkbookMutationAction.SetActiveSheet());
    ExecutorTestPlanSupport.PendingMutation setSelectedSheets =
        mutate(
            new SheetSelector.ByNames(List.of("Budget", "Budget Copy")),
            new WorkbookMutationAction.SetSelectedSheets());
    ExecutorTestPlanSupport.PendingMutation setSheetVisibility =
        mutate(
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN));
    ExecutorTestPlanSupport.PendingMutation setSheetProtection =
        mutate(
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.SetSheetProtection(protectionSettings()));
    ExecutorTestPlanSupport.PendingMutation clearSheetProtection =
        mutate(
            new SheetSelector.ByName("Budget"), new WorkbookMutationAction.ClearSheetProtection());

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
        mutate(
            new RangeSelector.ByRange("Budget", "A1:B2"), new WorkbookMutationAction.MergeCells());
    ExecutorTestPlanSupport.PendingMutation unmergeCells =
        mutate(
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new WorkbookMutationAction.UnmergeCells());
    ExecutorTestPlanSupport.PendingMutation setColumnWidth =
        mutate(
            new ColumnBandSelector.Span("Budget", 0, 1),
            new WorkbookMutationAction.SetColumnWidth(16.0));
    ExecutorTestPlanSupport.PendingMutation setRowHeight =
        mutate(
            new RowBandSelector.Span("Budget", 0, 1),
            new WorkbookMutationAction.SetRowHeight(28.5));
    ExecutorTestPlanSupport.PendingMutation insertRows =
        mutate(
            new RowBandSelector.Insertion("Budget", 1, 2), new WorkbookMutationAction.InsertRows());
    ExecutorTestPlanSupport.PendingMutation deleteRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.DeleteRows());
    ExecutorTestPlanSupport.PendingMutation shiftRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.ShiftRows(1));
    ExecutorTestPlanSupport.PendingMutation insertColumns =
        mutate(
            new ColumnBandSelector.Insertion("Budget", 1, 2),
            new WorkbookMutationAction.InsertColumns());
    ExecutorTestPlanSupport.PendingMutation deleteColumns =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new WorkbookMutationAction.DeleteColumns());
    ExecutorTestPlanSupport.PendingMutation shiftColumns =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new WorkbookMutationAction.ShiftColumns(-1));
    ExecutorTestPlanSupport.PendingMutation setRowVisibility =
        mutate(
            new RowBandSelector.Span("Budget", 1, 2),
            new WorkbookMutationAction.SetRowVisibility(true));
    ExecutorTestPlanSupport.PendingMutation setColumnVisibility =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new WorkbookMutationAction.SetColumnVisibility(false));
    ExecutorTestPlanSupport.PendingMutation groupRows =
        mutate(
            new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.GroupRows(true));
    ExecutorTestPlanSupport.PendingMutation ungroupRows =
        mutate(new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.UngroupRows());
    ExecutorTestPlanSupport.PendingMutation groupColumns =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new WorkbookMutationAction.GroupColumns(true));
    ExecutorTestPlanSupport.PendingMutation ungroupColumns =
        mutate(
            new ColumnBandSelector.Span("Budget", 1, 2),
            new WorkbookMutationAction.UngroupColumns());
    ExecutorTestPlanSupport.PendingMutation setSheetPane =
        mutate(
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1)));
    ExecutorTestPlanSupport.PendingMutation setSheetZoom =
        mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.SetSheetZoom(125));
    ExecutorTestPlanSupport.PendingMutation setSheetPresentation =
        mutate(
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.SetSheetPresentation(
                new SheetPresentationInput(
                    new SheetDisplayInput(false, false, false, true, true),
                    Optional.of(ColorInput.rgb("#112233")),
                    new SheetOutlineSummaryInput(false, false),
                    new SheetDefaultsInput(11, 18.5d),
                    List.of(
                        new IgnoredErrorInput(
                            "A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))));
    ExecutorTestPlanSupport.PendingMutation setPrintLayout =
        mutate(
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.SetPrintLayout(
                PrintLayoutInput.withDefaultSetup(
                    new PrintAreaInput.Range("A1:B12"),
                    ExcelPrintOrientation.LANDSCAPE,
                    new PrintScalingInput.Fit(1, 0),
                    new PrintTitleRowsInput.Band(0, 0),
                    new PrintTitleColumnsInput.Band(0, 0),
                    headerFooter("Budget", "", ""),
                    headerFooter("", "Page &P", ""))));
    ExecutorTestPlanSupport.PendingMutation clearPrintLayout =
        mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.ClearPrintLayout());

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
}
