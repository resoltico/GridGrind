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
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Workbook command and read-command translation coverage. */
class DefaultGridGrindRequestExecutorCommandTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void convertsWaveThreeOperationsIntoWorkbookCommands() {
    WorkbookCommand setHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/report"))));
    WorkbookCommand clearHyperlink =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.ClearHyperlink()));
    WorkbookCommand setComment =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(text("Review"), "GridGrind", false))));
    WorkbookCommand clearComment =
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"), new CellMutationAction.ClearComment()));
    WorkbookCommand setNamedRange =
        command(
            mutate(
                new StructuredMutationAction.SetNamedRange(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Budget", "B4"))));
    WorkbookCommand deleteNamedRange =
        command(
            mutate(
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                    "BudgetTotal", "Budget"),
                new StructuredMutationAction.DeleteNamedRange()));

    assertInstanceOf(WorkbookAnnotationCommand.SetHyperlink.class, setHyperlink);
    assertInstanceOf(WorkbookAnnotationCommand.ClearHyperlink.class, clearHyperlink);
    assertInstanceOf(WorkbookAnnotationCommand.SetComment.class, setComment);
    assertInstanceOf(WorkbookAnnotationCommand.ClearComment.class, clearComment);
    assertInstanceOf(WorkbookMetadataCommand.SetNamedRange.class, setNamedRange);
    assertInstanceOf(WorkbookMetadataCommand.DeleteNamedRange.class, deleteNamedRange);
    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"),
        cast(WorkbookAnnotationCommand.SetHyperlink.class, setHyperlink).target());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        cast(WorkbookAnnotationCommand.SetComment.class, setComment).comment());
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("Budget", "B4")),
        cast(WorkbookMetadataCommand.SetNamedRange.class, setNamedRange).definition());
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        cast(WorkbookMetadataCommand.DeleteNamedRange.class, deleteNamedRange).scope());
  }

  @Test
  void convertsRemainingWorkbookOperationsIntoWorkbookCommands() {
    assertInstanceOf(
        WorkbookSheetCommand.CreateSheet.class,
        command(
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.EnsureSheet())));
    assertInstanceOf(
        WorkbookSheetCommand.RenameSheet.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.RenameSheet("Summary"))));
    assertInstanceOf(
        WorkbookSheetCommand.DeleteSheet.class,
        command(
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.DeleteSheet())));
    assertInstanceOf(
        WorkbookSheetCommand.MoveSheet.class,
        command(
            mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.MoveSheet(0))));
    assertInstanceOf(
        WorkbookStructureCommand.MergeCells.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new WorkbookMutationAction.MergeCells())));
    assertInstanceOf(
        WorkbookStructureCommand.UnmergeCells.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B2"),
                new WorkbookMutationAction.UnmergeCells())));
    assertInstanceOf(
        WorkbookStructureCommand.SetColumnWidth.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 0, 1),
                new WorkbookMutationAction.SetColumnWidth(16.0))));
    assertInstanceOf(
        WorkbookStructureCommand.SetRowHeight.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 0, 1),
                new WorkbookMutationAction.SetRowHeight(28.5))));
    assertInstanceOf(
        WorkbookStructureCommand.InsertRows.class,
        command(
            mutate(
                new RowBandSelector.Insertion("Budget", 1, 2),
                new WorkbookMutationAction.InsertRows())));
    assertInstanceOf(
        WorkbookStructureCommand.DeleteRows.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.DeleteRows())));
    assertInstanceOf(
        WorkbookStructureCommand.ShiftRows.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.ShiftRows(1))));
    assertInstanceOf(
        WorkbookStructureCommand.InsertColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Insertion("Budget", 1, 2),
                new WorkbookMutationAction.InsertColumns())));
    assertInstanceOf(
        WorkbookStructureCommand.DeleteColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.DeleteColumns())));
    assertInstanceOf(
        WorkbookStructureCommand.ShiftColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.ShiftColumns(-1))));
    assertInstanceOf(
        WorkbookStructureCommand.SetRowVisibility.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.SetRowVisibility(true))));
    assertInstanceOf(
        WorkbookStructureCommand.SetColumnVisibility.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.SetColumnVisibility(true))));
    assertInstanceOf(
        WorkbookStructureCommand.GroupRows.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.GroupRows(false))));
    assertInstanceOf(
        WorkbookStructureCommand.UngroupRows.class,
        command(
            mutate(
                new RowBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.UngroupRows())));
    assertInstanceOf(
        WorkbookStructureCommand.GroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.GroupColumns(false))));
    assertInstanceOf(
        WorkbookStructureCommand.UngroupColumns.class,
        command(
            mutate(
                new ColumnBandSelector.Span("Budget", 1, 2),
                new WorkbookMutationAction.UngroupColumns())));
    assertInstanceOf(
        WorkbookLayoutCommand.SetSheetPane.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.SetSheetPane(new PaneInput.Frozen(1, 1, 1, 1)))));
    assertInstanceOf(
        WorkbookLayoutCommand.SetSheetZoom.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"), new WorkbookMutationAction.SetSheetZoom(125))));
    assertInstanceOf(
        WorkbookLayoutCommand.SetSheetPresentation.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.SetSheetPresentation(
                    new SheetPresentationInput(
                        new SheetDisplayInput(false, false, false, true, true),
                        java.util.Optional.of(ColorInput.rgb("#112233")),
                        new SheetOutlineSummaryInput(false, false),
                        new SheetDefaultsInput(11, 18.5d),
                        List.of(
                            new IgnoredErrorInput(
                                "A1:B2",
                                List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))))))));
    assertInstanceOf(
        WorkbookLayoutCommand.SetPrintLayout.class,
        command(
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
                        headerFooter("", "Page &P", ""))))));
    assertInstanceOf(
        WorkbookLayoutCommand.ClearPrintLayout.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.ClearPrintLayout())));
    assertInstanceOf(
        WorkbookCellCommand.SetCell.class,
        command(
            mutate(
                new CellSelector.ByAddress("Budget", "A1"),
                new CellMutationAction.SetCell(textCell("x")))));
    assertInstanceOf(
        WorkbookCellCommand.SetRange.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new CellMutationAction.SetRange(List.of(List.of(textCell("x")))))));
    assertInstanceOf(
        WorkbookCellCommand.ClearRange.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B1"),
                new CellMutationAction.ClearRange())));
    assertInstanceOf(
        WorkbookFormattingCommand.SetConditionalFormatting.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new StructuredMutationAction.SetConditionalFormatting(
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
        WorkbookFormattingCommand.ClearConditionalFormatting.class,
        command(
            mutate(
                new RangeSelector.AllOnSheet("Budget"),
                new StructuredMutationAction.ClearConditionalFormatting())));
    assertInstanceOf(
        WorkbookTabularCommand.SetAutofilter.class,
        command(
            mutate(
                new RangeSelector.ByRange("Budget", "A1:B4"),
                new StructuredMutationAction.SetAutofilter())));
    assertInstanceOf(
        WorkbookTabularCommand.ClearAutofilter.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new StructuredMutationAction.ClearAutofilter())));
    assertInstanceOf(
        WorkbookTabularCommand.SetTable.class,
        command(
            mutate(
                new StructuredMutationAction.SetTable(
                    TableInput.withDefaultMetadata(
                        "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None())))));
    assertInstanceOf(
        WorkbookTabularCommand.DeleteTable.class,
        command(
            mutate(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                new StructuredMutationAction.DeleteTable())));
    assertInstanceOf(
        WorkbookCellCommand.AppendRow.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"),
                new CellMutationAction.AppendRow(List.of(textCell("x"))))));
    assertInstanceOf(
        WorkbookLayoutCommand.AutoSizeColumns.class,
        command(
            mutate(
                new SheetSelector.ByName("Budget"), new WorkbookMutationAction.AutoSizeColumns())));

    WorkbookLayoutCommand.SetSheetPane setSheetPaneNone =
        cast(
            WorkbookLayoutCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetSheetPane(new PaneInput.None()))));
    WorkbookLayoutCommand.SetSheetPane setSheetPaneSplit =
        cast(
            WorkbookLayoutCommand.SetSheetPane.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetSheetPane(
                        new PaneInput.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT)))));
    WorkbookLayoutCommand.SetPrintLayout defaultPrintLayout =
        cast(
            WorkbookLayoutCommand.SetPrintLayout.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetPrintLayout(PrintLayoutInput.defaults()))));
    WorkbookLayoutCommand.SetSheetPresentation defaultSheetPresentation =
        cast(
            WorkbookLayoutCommand.SetSheetPresentation.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetSheetPresentation(
                        SheetPresentationInput.defaults()))));

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
    WorkbookSheetCommand.CopySheet copySheet =
        cast(
            WorkbookSheetCommand.CopySheet.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.CopySheet(
                        "Budget Copy", new SheetCopyPosition.AtIndex(1)))));
    WorkbookSheetCommand.SetActiveSheet setActiveSheet =
        cast(
            WorkbookSheetCommand.SetActiveSheet.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget Copy"),
                    new WorkbookMutationAction.SetActiveSheet())));
    WorkbookSheetCommand.SetSelectedSheets setSelectedSheets =
        cast(
            WorkbookSheetCommand.SetSelectedSheets.class,
            command(
                mutate(
                    new SheetSelector.ByNames(List.of("Budget", "Budget Copy")),
                    new WorkbookMutationAction.SetSelectedSheets())));
    WorkbookSheetCommand.SetSheetVisibility setSheetVisibility =
        cast(
            WorkbookSheetCommand.SetSheetVisibility.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN))));
    WorkbookSheetCommand.SetSheetProtection setSheetProtection =
        cast(
            WorkbookSheetCommand.SetSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.SetSheetProtection(protectionSettings()))));
    WorkbookSheetCommand.ClearSheetProtection clearSheetProtection =
        cast(
            WorkbookSheetCommand.ClearSheetProtection.class,
            command(
                mutate(
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.ClearSheetProtection())));

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
}
