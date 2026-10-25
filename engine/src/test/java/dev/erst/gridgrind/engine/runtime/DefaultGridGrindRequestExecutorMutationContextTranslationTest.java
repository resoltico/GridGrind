package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
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

    assertSheetOnlyContext(copySheet, exception, "Budget");
    assertSheetOnlyContext(setActiveSheet, exception, "Budget Copy");
    assertSheetOnlyContext(setSheetVisibility, exception, "Budget");
    assertSheetOnlyContext(setSheetProtection, exception, "Budget");
    assertSheetOnlyContext(clearSheetProtection, exception, "Budget");

    assertNull(formulaFor(setSelectedSheets, exception));
    assertNull(sheetNameFor(setSelectedSheets, exception));
    assertNull(addressFor(setSelectedSheets, exception));
    assertNull(rangeFor(setSelectedSheets, exception));

    assertNull(namedRangeNameFor(copySheet, exception));
    assertNull(namedRangeNameFor(setActiveSheet, exception));
    assertNull(namedRangeNameFor(setSelectedSheets, exception));
    assertNull(namedRangeNameFor(setSheetVisibility, exception));
    assertNull(namedRangeNameFor(setSheetProtection, exception));
    assertNull(namedRangeNameFor(clearSheetProtection, exception));
  }

  @Test
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

    assertRangeContext(mergeCells, exception, "Budget", "A1:B2");
    assertRangeContext(unmergeCells, exception, "Budget", "A1:B2");

    for (ExecutorTestPlanSupport.PendingMutation mutation :
        List.of(
            setColumnWidth,
            setRowHeight,
            insertRows,
            deleteRows,
            shiftRows,
            insertColumns,
            deleteColumns,
            shiftColumns,
            setRowVisibility,
            setColumnVisibility,
            groupRows,
            ungroupRows,
            groupColumns,
            ungroupColumns,
            setSheetPane,
            setSheetZoom,
            setSheetPresentation,
            setPrintLayout,
            clearPrintLayout)) {
      assertSheetOnlyContext(mutation, exception, "Budget");
    }
  }

  private static void assertSheetOnlyContext(
      ExecutorTestPlanSupport.PendingMutation mutation,
      Exception exception,
      String expectedSheetName) {
    assertNull(formulaFor(mutation, exception));
    assertEquals(expectedSheetName, sheetNameFor(mutation, exception));
    assertNull(addressFor(mutation, exception));
    assertNull(rangeFor(mutation, exception));
  }

  private static void assertRangeContext(
      ExecutorTestPlanSupport.PendingMutation mutation,
      Exception exception,
      String expectedSheetName,
      String expectedRange) {
    assertNull(formulaFor(mutation, exception));
    assertEquals(expectedSheetName, sheetNameFor(mutation, exception));
    assertNull(addressFor(mutation, exception));
    assertEquals(expectedRange, rangeFor(mutation, exception));
  }
}
