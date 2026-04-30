package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Authoring-metadata and fallback context extraction coverage. */
class DefaultGridGrindRequestExecutorAuthoringContextTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void extractsContextForAuthoringMetadataAndNamedRangeOperations() {
    RuntimeException exception = new RuntimeException("test");
    assertWriteContext(
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new CellMutationAction.SetHyperlink(
                new HyperlinkTarget.Url("https://example.com/report"))),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(new CellSelector.ByAddress("Budget", "A1"), new CellMutationAction.ClearHyperlink()),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new CellMutationAction.SetComment(
                CommentInput.plain(text("Review"), "GridGrind", false))),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(new CellSelector.ByAddress("Budget", "A1"), new CellMutationAction.ClearComment()),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new CellMutationAction.ApplyStyle(
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
            new StructuredMutationAction.SetDataValidation(
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
        mutate(
            new RangeSelector.AllOnSheet("Budget"),
            new StructuredMutationAction.ClearDataValidations()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new SheetSelector.ByName("Budget"),
            new StructuredMutationAction.SetConditionalFormatting(
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
            new StructuredMutationAction.ClearConditionalFormatting()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new RangeSelector.ByRange("Budget", "A1:C4"),
            new StructuredMutationAction.SetAutofilter()),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        mutate(new SheetSelector.ByName("Budget"), new StructuredMutationAction.ClearAutofilter()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None()))),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        mutate(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
            new StructuredMutationAction.DeleteTable()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        mutate(
            new StructuredMutationAction.SetNamedRange(
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
            new StructuredMutationAction.DeleteNamedRange()),
        exception,
        null,
        null,
        null,
        "BudgetTotal");
    assertWriteContext(
        mutate(
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                "LocalItem", "Budget"),
            new StructuredMutationAction.DeleteNamedRange()),
        exception,
        "Budget",
        null,
        null,
        "LocalItem");
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.EnsureSheet()),
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
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.EnsureSheet()),
            invalidFormula));
    assertEquals(
        "BAD!",
        addressFor(
            mutate(
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                    "BudgetTotal"),
                new StructuredMutationAction.DeleteNamedRange()),
            invalidAddress));

    assertEquals(
        "C1:D2",
        rangeFor(
            mutate(
                new RangeSelector.ByRange("Budget", "C1:D2"),
                new CellMutationAction.SetRange(List.of(List.of(textCell("x"))))),
            invalidFormula));
    assertEquals(
        "E1:E2",
        rangeFor(
            mutate(
                new RangeSelector.ByRange("Budget", "E1:E2"), new CellMutationAction.ClearRange()),
            invalidFormula));
    assertEquals(
        "B2:B5",
        rangeFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new StructuredMutationAction.SetConditionalFormatting(
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
                new StructuredMutationAction.SetConditionalFormatting(
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
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.ClearHyperlink()),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(text("Review"), "GridGrind", false))),
            invalidFormula));
    assertNull(
        rangeFor(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"), new CellMutationAction.ClearComment()),
            invalidFormula));
    assertEquals(
        "A1:",
        rangeFor(
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.EnsureSheet()),
            invalidRange));

    List<ExecutorTestPlanSupport.PendingMutation> operationsWithoutNamedRanges =
        List.of(
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.EnsureSheet()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.RenameSheet("Summary")),
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.DeleteSheet()),
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.MoveSheet(0)),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new WorkbookMutationAction.MergeCells()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new WorkbookMutationAction.UnmergeCells()),
            mutate(
                new ColumnBandSelector.Span("Budget", 0, 1),
                new WorkbookMutationAction.SetColumnWidth(16.0)),
            mutate(
                new RowBandSelector.Span("Budget", 0, 1),
                new WorkbookMutationAction.SetRowHeight(28.5)),
            mutate(
                new RowBandSelector.Insertion("Budget", 1, 2),
                new WorkbookMutationAction.InsertRows()),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.DeleteRows()),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.ShiftRows(1)),
            mutate(
                new ColumnBandSelector.Insertion("Budget", 1, 2),
                new WorkbookMutationAction.InsertColumns()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.DeleteColumns()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.ShiftColumns(-1)),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.SetRowVisibility(true)),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.SetColumnVisibility(false)),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.GroupRows(true)),
            mutate(
                new RowBandSelector.Span("Budget", 1, 2), new WorkbookMutationAction.UngroupRows()),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.GroupColumns(true)),
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.UngroupColumns()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
            mutate(
                new SheetSelector.ByName("Budget"), new WorkbookMutationAction.SetSheetZoom(125)),
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
                        headerFooter("", "Page &P", "")))),
            mutate(
                new SheetSelector.ByName("Budget"), new WorkbookMutationAction.ClearPrintLayout()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetCell(textCell("x"))),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new CellMutationAction.SetRange(List.of(List.of(textCell("x"))))),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"), new CellMutationAction.ClearRange()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.ClearHyperlink()),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(text("Review"), "GridGrind", false))),
            mutate(
                new CellSelector.ByAddress("Budget", "A1"), new CellMutationAction.ClearComment()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new CellMutationAction.ApplyStyle(
                    new CellStyleInput(
                        null,
                        null,
                        new CellFontInput(true, null, null, null, null, null, null),
                        null,
                        null,
                        null))),
            mutate(
                new RangeSelector.ByRange("Budget", "B2:B5"),
                new StructuredMutationAction.SetDataValidation(
                    new DataValidationInput(
                        new DataValidationRuleInput.TextLength(
                            ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                        true,
                        false,
                        null,
                        null))),
            mutate(
                new RangeSelector.AllOnSheet("Budget"),
                new StructuredMutationAction.ClearDataValidations()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new StructuredMutationAction.SetConditionalFormatting(
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
                new StructuredMutationAction.ClearConditionalFormatting()),
            mutate(
                new RangeSelector.ByRange("Budget", "A1:C4"),
                new StructuredMutationAction.SetAutofilter()),
            mutate(
                new SheetSelector.ByName("Budget"), new StructuredMutationAction.ClearAutofilter()),
            mutate(
                new StructuredMutationAction.SetTable(
                    TableInput.withDefaultMetadata(
                        "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None()))),
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new StructuredMutationAction.DeleteTable()),
            mutate(
                new SheetSelector.ByName("Budget"),
                new CellMutationAction.AppendRow(List.of(textCell("x")))),
            mutate(
                new SheetSelector.ByName("Budget"), new WorkbookMutationAction.AutoSizeColumns()));

    for (ExecutorTestPlanSupport.PendingMutation operation : operationsWithoutNamedRanges) {
      assertNull(namedRangeNameFor(operation, invalidFormula));
    }

    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1))),
            missingNamedRange));
  }
}
